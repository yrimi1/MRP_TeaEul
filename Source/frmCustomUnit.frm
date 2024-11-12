VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmCustomUnit 
   Caption         =   "거래처별 단가관리"
   ClientHeight    =   7470
   ClientLeft      =   2820
   ClientTop       =   3885
   ClientWidth     =   10920
   Icon            =   "frmCustomUnit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   10920
   Begin VSFlex7LCtl.VSFlexGrid grdCustom 
      Height          =   5640
      Left            =   15
      TabIndex        =   29
      Top             =   960
      Width           =   3630
      _cx             =   6403
      _cy             =   9948
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
   Begin Threed.SSPanel pnlBoard 
      Height          =   915
      Left            =   3735
      TabIndex        =   13
      Top             =   30
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   1614
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   3885
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   10
         ToolTipText     =   "자료 취소"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   5475
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   11
         ToolTipText     =   "자료 수정"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   6270
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   12
         ToolTipText     =   "자료 삭제"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   4680
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   0
         ToolTipText     =   "자료 추가"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   3090
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   9
         ToolTipText     =   "자료 저장"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   510
         Left            =   75
         TabIndex        =   30
         Top             =   330
         Visible         =   0   'False
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   900
         _Version        =   196609
         BackColor       =   65535
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   915
      Left            =   30
      TabIndex        =   14
      Top             =   30
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   1614
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtSearch 
         Height          =   300
         IMEMode         =   10  '한글 
         Left            =   75
         TabIndex        =   15
         Top             =   465
         Width           =   2025
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   120
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "상호명 검색"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdAll 
         Height          =   330
         Left            =   2160
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   450
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   582
         _Version        =   196609
         MousePointer    =   99
         CaptionStyle    =   1
         PictureAnimationEnabled=   0   'False
         Alignment       =   6
         PictureAlignment=   0
         BevelWidth      =   1
         ShapeSize       =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   3090
      Left            =   3705
      TabIndex        =   18
      Top             =   975
      Width           =   7140
      _cx             =   12594
      _cy             =   5450
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
   Begin Threed.SSPanel pnlEdit 
      Height          =   2505
      Left            =   3720
      TabIndex        =   19
      Top             =   4080
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   4419
      _Version        =   196609
      Enabled         =   0   'False
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboWork 
         Height          =   300
         Left            =   1380
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   1140
         Width           =   2235
      End
      Begin VB.ComboBox cboWidth 
         Height          =   300
         IMEMode         =   2  '입력 상태 해제
         Left            =   1380
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   780
         Width           =   2235
      End
      Begin VB.TextBox txtCustom 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   75
         Width           =   3285
      End
      Begin VB.TextBox txtETC 
         Height          =   300
         Left            =   1380
         TabIndex        =   8
         Top             =   2145
         Width           =   4245
      End
      Begin MRPPlus2.WizText txtArticle 
         Height          =   300
         Left            =   1380
         TabIndex        =   2
         Top             =   405
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MRPPlus2.WizText txtCode 
         Height          =   300
         Left            =   6225
         TabIndex        =   20
         Top             =   60
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MaxLength       =   4
         Text            =   "123"
         BackColor       =   12648384
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   4920
         TabIndex        =   21
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "거래처 코드"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   90
         TabIndex        =   22
         Top             =   420
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "품          명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   90
         TabIndex        =   23
         Top             =   780
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "가   공   폭"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   7
         Left            =   90
         TabIndex        =   24
         Top             =   1815
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "가 공  단 가"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   5
         Left            =   90
         TabIndex        =   25
         Top             =   1140
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "가   공   명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   6
         Left            =   90
         TabIndex        =   26
         Top             =   1470
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "축율 + Loss"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   8
         Left            =   90
         TabIndex        =   27
         Top             =   2145
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "비        고"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   3630
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   405
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         PictureFrames   =   1
         Enabled         =   0   'False
         Picture         =   "frmCustomUnit.frx":000C
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   9
         Left            =   90
         TabIndex        =   28
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "상   호    명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtChunkRate 
         Height          =   300
         Left            =   1380
         TabIndex        =   6
         Top             =   1470
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin MRPPlus2.WizText txtUnitPrice 
         Height          =   300
         Left            =   1380
         TabIndex        =   7
         Top             =   1800
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   9210
      TabIndex        =   32
      Top             =   6630
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Label lblCount 
      Caption         =   "검색건수 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   31
      Top             =   6645
      Width           =   3600
   End
End
Attribute VB_Name = "frmCustomUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIMIT_WIDTH As Integer = 3140
Private Const LIMIT_ROW = 16

Private m_sOperate     As String * 1
Private m_bSortForward As Boolean









Private Sub cboWidth_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboWork_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_ARTICLE, , True, txtArticle)
    End If
    
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11040, 7875

    Call SetOperate(Me)

    cmdAll.Picture = LoadResPicture("ALL", vbResIcon)
    
    pnlCaption(9).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(3).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(4).Picture = LoadResPicture("BASIC", vbResIcon)
    pnlCaption(5).Picture = LoadResPicture("BASIC", vbResIcon)

    Call InitGrid
    Call MakeCombo

    Call FillGrid
End Sub

Private Sub Form_Activate()
    txtSearch.SetFocus
End Sub


Private Sub MakeCombo()
    Call MakeCodeCombo(cboWork, CD_WORK)        ' 가공 구분
    Call SetStuffWidth
    
    
''    ' 화폐단위
''    With cboPriceClss
''        .AddItem "1. Dollar ($)"
''        .AddItem "2. Won (\)"
''
''        .ListIndex = 0
''    End With
''
''    ' 길이단위
''    With cboUnitClss
''        .AddItem "1. Yard"
''        .AddItem "2. Meter"
''
''        .ListIndex = 0
''    End With
    
End Sub


Private Sub grdCustom_RowColChange()
    With grdCustom
        If .Rows = .FixedRows Then Exit Sub
        
        txtCustom.Text = .TextMatrix(.Row, 2)
        txtCode.Text = .TextMatrix(.Row, 1)
    
    End With
    
    Call ShowData
    
End Sub

Private Sub SetStuffWidth()
    Dim oCode As Pluslib2.CCode
    Dim rs    As ADODB.Recordset
    Dim II%
    
    On Error GoTo ErrHandler

    Set oCode = New Pluslib2.CCode
    oCode.Connection = g_adoCon

    Set rs = oCode.GetStuffWidth
    Set oCode = Nothing
    II = 0
    cboWidth.Clear
    If Not rs Is Nothing Then
        If Not rs.BOF Then
           rs.MoveFirst
           Do Until rs.EOF
            cboWidth.AddItem Trim$(rs(0))
            cboWidth.ItemData(II) = val(rs(1))
            II = II + 1
            rs.MoveNext
           Loop
        End If
    End If

    rs.Close
    Set rs = Nothing
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCode = Nothing

    Err.Raise Err.Number, "frmCustomUnit.SetStuffWidth", Err.Description, Err.HelpFile, Err.HelpContext

End Sub






Private Sub txtArticle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call ReturnRef(LG_ARTICLE, , False, txtArticle)
        If Len(txtArticle.Tag) < 0 Then
            txtArticle.SetFocus
            'Call MoveFocus(KeyCode)
        Else
            cboWidth.SetFocus
        End If
    End If
End Sub

Private Sub txtArticle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnRef(LG_ARTICLE, , False, txtArticle)
        If Len(txtArticle.Tag) < 0 Then
            txtArticle.SetFocus
            'Call MoveFocus(KeyCode)
        End If
    End If
End Sub

Private Sub txtETC_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)

End Sub

Private Sub txtSearch_Change()
    Dim i%, iCount%, iNowRow%

    On Error GoTo ErrHandler

    If Len(Trim(txtSearch)) > 0 Then
        With grdCustom
            .Redraw = flexRDNone

            For i = .FixedRows To .Rows - .FixedRows
                If InStr(UCase(.TextArray(i * .Cols + 2)), UCase(txtSearch)) = 0 Then
                    .RowHidden(i) = True
                    iCount = iCount + 1
                Else
                    .RowHidden(i) = False
                    iNowRow = i
                End If
            Next i

            If iNowRow > .FixedRows Then
                .Row = iNowRow
                .Col = .FixedCols
                .ColSel = .Cols - 1
            End If

            .Redraw = flexRDDirect
        End With
    Else
        Call cmdAll_Click
    End If

    If iCount > 0 Then
        cmdAll.Visible = True
    Else
        cmdAll.Visible = False
    End If

    Call ChangeScroll

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "txtSearch.Change", Err.Description)
End Sub

Private Sub ChangeScroll()
    Dim lRows As Long

    On Error GoTo ErrHandler
    
    lRows = GetVisibleVSGridRowCount(grdData)

    With grdCustom
        .Redraw = flexRDNone
        If .Rows > LIMIT_ROW Then
            .ColWidth(2) = LIMIT_WIDTH - 240
        Else
            .ColWidth(2) = LIMIT_WIDTH
        End If
        .Redraw = flexRDDirect
    End With
    
    If lRows = 0 Then
        cmdOperate(ID_UPDATE).Enabled = False
        cmdOperate(ID_DELETE).Enabled = False
    Else
        cmdOperate(ID_UPDATE).Enabled = True
        cmdOperate(ID_DELETE).Enabled = True
    End If
    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "CustomUnit.ChangeScroll", Err.Description)

End Sub


Private Sub cmdAll_Click()
    Dim i%

    With grdCustom
        .Redraw = flexRDNone

        For i = .FixedRows To .Rows - .FixedRows
            .RowHidden(i) = False
        Next i

        .Redraw = flexRDDirect
    End With

    txtSearch.Text = ""
    cmdAll.Visible = False
End Sub


Private Sub grdData_DblClick()
    With grdData
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub

        Call cmdOperate_Click(ID_UPDATE)
        txtChunkRate.SetFocus
    End With
End Sub


Private Sub grdData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdOperate_Click(ID_UPDATE)
End Sub


Private Sub grdData_RowColChange()
    Call ShowDataDetail
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim bResult As Boolean

    On Error GoTo ErrHandler

    Select Case Index
    Case ID_ADDNEW
        m_sOperate = ID_ADDNEW
        Call ChangeMode(Me, False)
        Call ClearData
        Call ChangeEditMode
        
        pnlMsg.Caption = LoadResString(302)

        pnlEdit.Enabled = True

    Case ID_UPDATE '[2] 수정
        m_sOperate = ID_UPDATE
        Call ChangeMode(Me, False)
        Call ChangeEditMode
        
        pnlMsg.Caption = LoadResString(303)
        pnlEdit.Enabled = True
        
        txtChunkRate.SetFocus

    Case ID_DELETE '[3] 삭제
        If grdData.Rows = grdData.FixedRows Then Exit Sub

        If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
            m_sOperate = ID_DELETE

            If DelData() Then Call ShowData
        End If
        
    Case ID_SAVE  '[4] 저장
        If SaveData() Then
            Call ShowData
            Call ChangeMode(Me, True)

            m_sOperate = ""
            
            pnlEdit.Enabled = False
        End If
        grdData.SetFocus
        
    Case ID_CANCEL
        m_sOperate = ""
        If grdData.Rows > 1 Then
            Call ShowData
        Else
            Call ClearData
        End If
        Call ChangeMode(Me, True)
         
        pnlEdit.Enabled = False

        grdData.SetFocus
    End Select

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub


Private Sub ChangeEditMode()

    If m_sOperate = ID_ADDNEW Then
        txtCode.Locked = False
        txtCustom.Locked = False
'        cboTrade.Locked = False
        txtArticle.Locked = False
        '.Locked = False
'        txtWork.Locked = False
'        cboPriceClss.Locked = False
'        cboUnitClss.Locked = False
'        txtPrice.Locked = False
    ElseIf m_sOperate = ID_UPDATE Then
        txtCode.Locked = True
        txtCustom.Locked = True
'        cboTrade.Locked = True
        txtArticle.Locked = True
'        txtColor.Locked = True
'        txtWork.Locked = True
'        cboPriceClss.Locked = True
'        cboUnitClss.Locked = True
'        txtPrice.Locked = False
    End If

End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub ClearData()
    txtArticle = ""
    txtArticle.Tag = 0
    cboWidth.ListIndex = 0
    cboWork.ListIndex = 0
    txtChunkRate = ""
    txtUnitPrice = ""
    txtETC.Text = ""
    cmdFind(1).Enabled = True

End Sub



Private Sub InitGrid()
    
    With grdCustom
        .Cols = 3
        Call SetVSFlexGrid(grdCustom)

        .Redraw = flexRDNone
        .Rows = 1

        .TextArray(0) = ""
        .TextArray(1) = "코드":         .ColWidth(1) = 800:     .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "거래처 명":    .ColWidth(2) = 1000:    .ColAlignment(2) = flexAlignLeftCenter

        .Redraw = flexRDDirect
    
    End With


    With grdData
        .Cols = 11
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone
        .Rows = 1

        .TextArray(0) = ""
        .TextArray(1) = "거래처코드":   .ColWidth(1) = 0:       .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "품 명":        .ColWidth(2) = 2000:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "품명코드":     .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "가공폭":       .ColWidth(4) = 800:     .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "가공폭코드":   .ColWidth(5) = 0:       .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "가공구분":     .ColWidth(6) = 1300:    .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "가공구분코드": .ColWidth(7) = 0:       .ColAlignment(7) = flexAlignLeftCenter
        .TextArray(8) = "축율":         .ColWidth(8) = 600:     .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "단가":         .ColWidth(9) = 600:     .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "비고":        .ColWidth(10) = 1000:   .ColAlignment(10) = flexAlignLeftCenter

        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub FillGrid()
    Dim oCustom As Pluslib2.CCustom
    Dim rs      As ADODB.Recordset
    Dim lNowRow&

    On Error GoTo ErrHandler

    Set oCustom = New Pluslib2.CCustom
    oCustom.Connection = g_adoCon
    
    Set rs = oCustom.GetCustom()
    
    Set oCustom = Nothing
    
    With grdCustom
        .Redraw = flexRDNone

        lNowRow = IIf(.Row > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows

        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs!CustomID & vbTab & rs!kCustom
            
            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing

     
        If .Rows > .FixedRows Then
           .HighLight = flexHighlightAlways
           .Row = 1
           
           .Col = .FixedCols
           .ColSel = .Cols - 1

            lblCount.Caption = LoadResString(250) & grdCustom.Rows - 1 & " 건"
            'Call ShowData
        Else
            .HighLight = flexHighlightNever

            Call ClearData
        End If

        

        .Redraw = flexRDDirect
    End With

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCustom = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub


Private Sub ShowData()
    Dim oCustom As Pluslib2.CCustom
    Dim rs      As ADODB.Recordset
    Dim lNowRow&
    Dim sCustomID$

    On Error GoTo ErrHandler

    Set oCustom = New Pluslib2.CCustom
    oCustom.Connection = g_adoCon
    
    sCustomID = grdCustom.TextMatrix(grdCustom.Row, 1)
    
    Set rs = oCustom.GetCustomUnit(sCustomID)
    Set oCustom = Nothing
    
    With grdData
        .Redraw = flexRDNone

        lNowRow = IIf(.Row > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows

        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs!CustomID & vbTab & CheckNull(rs!Article) & vbTab & CheckNull(rs!ArticleID) & vbTab & _
                CheckNull(rs!StuffWidth) & vbTab & CheckNull(rs!StuffWidthID) & vbTab & CheckNull(rs!WorkName) & vbTab & _
                CheckNull(rs!WorkID) & vbTab & Format(CheckNull(rs!ChunkRate), "###.00") & vbTab & CheckNull(rs!UnitPrice) & vbTab & _
                CheckNull(rs!ETC)
        
            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
           .HighLight = flexHighlightAlways
           .Row = IIf(.Rows > lNowRow, lNowRow, .Rows - 1)
           .TopRow = lNowRow

           .Col = .FixedCols
           .ColSel = .Cols - 1

        Else
            .HighLight = flexHighlightNever

            Call ClearData
        End If

        

        .Redraw = flexRDDirect
    End With

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCustom = Nothing
    

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub


Private Sub ShowDataDetail()

    If grdData.Rows = grdData.FixedRows Then Exit Sub

    With grdData
        txtCode = grdCustom.TextMatrix(grdCustom.Row, 1)
        txtCustom = grdCustom.TextMatrix(grdCustom.Row, 2)
        txtArticle = .TextMatrix(.Row, 2)
        txtArticle.Tag = .TextMatrix(.Row, 3)
        txtChunkRate = .ValueMatrix(.Row, 8)
        txtUnitPrice = .ValueMatrix(.Row, 9)
        txtETC = .TextMatrix(.Row, 10)
        cboWidth.ListIndex = FindComboBox(cboWidth, .ValueMatrix(.Row, 5))    '가공폭
        cboWork.ListIndex = FindComboBox(cboWork, .ValueMatrix(.Row, 7))      '가공구분
        
    End With
End Sub

Private Function DelData() As Boolean
    Dim oCustom   As Pluslib2.CCustom
    Dim tUnit     As Pluslib2.TCustomUnit

    On Error GoTo ErrHandler
    
    Set oCustom = New Pluslib2.CCustom
    oCustom.Connection = g_adoCon
    oCustom.UserName = g_sUserName
    
    
    
    With tUnit
        .sCustomID = txtCode
        .sArticleID = txtArticle.Tag
        .sStuffWidthID = Format(cboWidth.ItemData(cboWidth.ListIndex), "0#")
        .sWorkID = Format(cboWork.ItemData(cboWork.ListIndex), "000#")
    End With
    
    DelData = oCustom.DeleteCustomUnit(tUnit)

    Set oCustom = Nothing

    Exit Function

ErrHandler:
    Set oCustom = Nothing
    
    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Function

Private Function SaveData() As Boolean
    Dim oCustom   As Pluslib2.CCustom
    Dim tUnit     As Pluslib2.TCustomUnit

    On Error GoTo ErrHandler
    
    Set oCustom = New Pluslib2.CCustom
    oCustom.Connection = g_adoCon
    oCustom.UserName = g_sUserName
    
    If Len(txtArticle.Tag) <> 4 Then
        MsgBox ("품명코드를 다시 확인 하십시오")
        SaveData = False
        Exit Function
    End If
    
    With tUnit
        .sCustomID = txtCode
        .sArticleID = txtArticle.Tag
        .sStuffWidthID = Format(cboWidth.ItemData(cboWidth.ListIndex), "0#")
        .sWorkID = Format(cboWork.ItemData(cboWork.ListIndex), "000#")
        .nChunkRate = val(txtChunkRate)
        .nUnitPrice = Format(txtUnitPrice, "##0.00")
        .sETC = Trim(txtETC)
    
    End With
    
    SaveData = oCustom.AddNewCustomUnit(tUnit)

    Set oCustom = Nothing

    Exit Function

ErrHandler:
    Set oCustom = Nothing
    
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Function

