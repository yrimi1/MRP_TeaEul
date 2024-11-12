VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffINView 
   Caption         =   "생지 입고 조회"
   ClientHeight    =   9255
   ClientLeft      =   1770
   ClientTop       =   3465
   ClientWidth     =   15180
   Icon            =   "frmStuffINView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15180
   Begin VB.ComboBox cboOrderID 
      Height          =   300
      Left            =   12270
      Style           =   2  '드롭다운 목록
      TabIndex        =   32
      Top             =   390
      Width           =   1965
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   9150
      TabIndex        =   29
      Top             =   60
      Width           =   1695
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   7755
      Left            =   30
      TabIndex        =   22
      Top             =   780
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   13679
      _Version        =   196609
      Caption         =   "SSPanel3"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdShrink 
         Caption         =   "축소"
         Height          =   345
         Index           =   1
         Left            =   3090
         TabIndex        =   28
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdShrink 
         Caption         =   "확장"
         Height          =   345
         Index           =   0
         Left            =   2310
         TabIndex        =   27
         Top             =   30
         Width           =   765
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "오더별"
         Height          =   345
         Index           =   0
         Left            =   30
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   30
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "거래처별"
         Height          =   345
         Index           =   1
         Left            =   1080
         Style           =   1  '그래픽
         TabIndex        =   23
         Top             =   30
         Width           =   990
      End
      Begin VSFlex7LCtl.VSFlexGrid grdGroup 
         Height          =   6960
         Left            =   30
         TabIndex        =   25
         Top             =   390
         Width           =   15030
         _cx             =   26511
         _cy             =   12277
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
      Begin VSFlex7LCtl.VSFlexGrid grdTotal 
         Height          =   330
         Left            =   30
         TabIndex        =   26
         Top             =   7380
         Width           =   15030
         _cx             =   26511
         _cy             =   582
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
   Begin VB.TextBox txtArticle 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5490
      TabIndex        =   8
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   5490
      TabIndex        =   7
      Top             =   30
      Width           =   1935
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   1530
      MousePointer    =   99  '사용자 정의
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   390
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전일"
      Height          =   315
      Index           =   0
      Left            =   2190
      MousePointer    =   99  '사용자 정의
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   390
      Width           =   615
   End
   Begin VB.ComboBox CboStuffClss2 
      Height          =   300
      Left            =   12270
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   30
      Width           =   1965
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   720
      Left            =   14310
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   1
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   810
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13500
      TabIndex        =   0
      Top             =   8550
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Left            =   10950
      TabIndex        =   3
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "입고구분"
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   1065
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   315
      Index           =   0
      Left            =   7470
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   30
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   9
      Left            =   4170
      TabIndex        =   10
      Top             =   30
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
         TabIndex        =   11
         Top             =   60
         Width           =   975
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   11
      Left            =   4170
      TabIndex        =   12
      Top             =   360
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "품     명"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   1065
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   315
      Index           =   2
      Left            =   7470
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   390
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   196609
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   2850
      TabIndex        =   15
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   71368705
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   2850
      TabIndex        =   16
      Top             =   390
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   71368705
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   1530
      TabIndex        =   17
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "입고 일자"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   18
         Top             =   60
         Value           =   1  '확인
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   675
      Left            =   30
      TabIndex        =   19
      Top             =   30
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1191
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   120
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   3
         Left            =   90
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   390
         Width           =   1140
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   7830
      TabIndex        =   30
      Top             =   60
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
         Caption         =   "관리번호"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   31
         Top             =   60
         Width           =   1185
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   315
      Left            =   10950
      TabIndex        =   33
      Top             =   390
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "확정구분"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   34
         Top             =   60
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmStuffINView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_iFlag As String * 1
Dim m_bGroupClss As Boolean     '거래처별, 오더별 Grid 구분

Private Const LIMIT_WIDTH1 = 1640
Private Const LIMIT_WIDTH2 = 2100
Private Const LIMIT_WIDTH3 = 560
Private Const LIMIT_WIDTH4 = 2000
Private Const LIMIT_ROW1 = 11
Private Const LIMIT_ROW2 = 28
Private Const LIMIT_ROW3 = 9
Private m_bSortForward As Boolean
Private m_StuffDate As String, m_StuffClss As String, m_StuffSeq As Integer

Private Sub GridCollapse(oFlex As VSFlexGrid, Row As Integer)
    With oFlex
        If Row < .FixedRows Then Exit Sub

        If .IsCollapsed(Row) = flexOutlineCollapsed Then
            .IsCollapsed(Row) = flexOutlineExpanded
        Else
            .IsCollapsed(Row) = flexOutlineCollapsed
        End If
    End With
End Sub



Private Sub cmdShrink_Click(Index As Integer)
    If Index = 0 Then
        Call SetGrdShrink(grdGroup, OM_EXPAND)
    Else
        Call SetGrdShrink(grdGroup, OM_REDUCE)
    End If
    
''    Dim II As Integer
''    Dim nRows As String, sRows_var As Variant
''
''    nRows = ""
''    With grdGroup
''        Select Case Index
''            Case 0
''                For II = .FixedRows To .Rows - 1
''                    If .IsCollapsed(II) = flexOutlineCollapsed Then
''                        nRows = nRows & "," & II
''                    End If
''                Next II
''            Case 1
''                For II = .Rows - 1 To .FixedRows Step -1
''                    If .IsCollapsed(II) = flexOutlineExpanded And .IsSubtotal(II) Then
''                        nRows = nRows & "," & II
''                    End If
''                Next II
''        End Select
''    End With
''
''    nRows = Mid(nRows, 2)
''
''    sRows_var = Split(nRows, ",")
''
''    For II = 0 To UBound(sRows_var)
''        Call GridCollapse(grdGroup, val(sRows_var(II)))
''    Next II
End Sub

Private Sub dtpDate_KeyPress(Index As Integer, KeyAscii As Integer)
    Call MoveFocus(KeyAscii)
End Sub

Private Sub Form_Load()
    Dim i%
    
    PlusMDI.pnlMenu.Visible = False
    
    Me.Move 0, 0, 15300, 9660

    Call InitGroup
    
    Call SetOperate(Me)
    
    '----- 검색용 입고구분 설정
    With CboStuffClss2
        .AddItem "1.생지"
        .ItemData(0) = 1
        .AddItem "3.반품 생지"
        .ItemData(1) = 3
        .ListIndex = 0
    End With
    
    '----- 확정구분
    With cboOrderID
        .AddItem "수주확정"
        .AddItem "수주미확정"
        .ListIndex = 0
    End With
    
    '---- 날짜 설정
    For i = 0 To 1
        dtpDate(i) = Now
    Next i
    
    '--- find 컨트롤 icon설정
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    cmdFind(2).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(2).MouseIcon = LoadResPicture("POINTER", vbResCursor)

    cmdFind(0).Enabled = False
    cmdFind(2).Enabled = False
    
    txtCustom(1).Enabled = False
    txtArticle.Enabled = False
    txtSearch(3).Enabled = False
    CboStuffClss2.Enabled = False
    cboOrderID.Enabled = False
    

    m_iFlag = ID_ADDNEW
    
    '---- 오더별 데이터 나타내기
    m_bGroupClss = True
    Call FillGridGroup(m_bGroupClss)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Call SaveSetting(LoadResString(100), Me.Name, "Order", IIf(chkSearch(0) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "Custom", IIf(chkSearch(1) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "Article", IIf(chkSearch(2) = vbChecked, "1", "0"))
End Sub

''Sub SetKeyEdit(ByVal dEdit As Boolean)
''    dtpDate(2).Enabled = dEdit
''    CboStuffClss.Enabled = dEdit
''    txtStuffSeq.Enabled = dEdit
''End Sub

Private Sub chkSearch_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
       Case 0    '관리번호
            If chkSearch(0) Then
                txtSearch(3).Enabled = True
                txtSearch(3).SetFocus
            Else
                txtSearch(3).Enabled = False
                txtSearch(3).Text = ""
            End If
        Case 1    '거래처
            If chkSearch(1) = vbChecked Then
                txtCustom(1).Enabled = True
                txtCustom(1).SetFocus
                cmdFind(0).Enabled = True
            Else
                txtCustom(1).Enabled = False
                cmdFind(0).Enabled = False
                txtCustom(1).Tag = ""
            End If
        Case 2    '품명
            If chkSearch(2) = vbChecked Then
                txtArticle.Enabled = True
                txtArticle.SetFocus
                cmdFind(2).Enabled = True
            Else
                txtArticle.Enabled = False
                txtArticle.Tag = ""
                cmdSearch.SetFocus
                cmdFind(2).Enabled = False
            End If
        Case 3     '입고일자 Term
            If chkSearch(3) = vbChecked Then
                dtpDate(0).Enabled = True
                dtpDate(1).Enabled = True
            Else
                dtpDate(0).Enabled = False
                dtpDate(1).Enabled = False
            End If
        Case 4     '입고구분
            If chkSearch(Index) = vbChecked Then
                CboStuffClss2.Enabled = True
            Else
                CboStuffClss2.Enabled = False
            End If
        Case 5     '확정구분
            If chkSearch(5) = vbChecked Then
                cboOrderID.Enabled = True
            Else
                cboOrderID.Enabled = False
            End If
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


''Private Sub ChangeScroll(Index As Integer)
''    Select Case Index
''
''    Case 1
''        With grdGroup
''            If m_bGroupClss Then
''                If .Rows > LIMIT_ROW2 Then
''                    .ColWidth(5) = LIMIT_WIDTH2 - 240
''                Else
''                    .ColWidth(5) = LIMIT_WIDTH2
''                End If
''            Else
''                If .Rows > LIMIT_ROW2 Then
''                    .ColWidth(7) = LIMIT_WIDTH4 - 240
''                Else
''                    .ColWidth(7) = LIMIT_WIDTH4
''                End If
''
''            End If
''        End With
''    End Select
''End Sub

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0                '[1] 거래처 코드
            Call ReturnCode(LG_CUSTOM, , False, txtCustom(1))
''        Case 1                '[2] 거래처 코드
''            Call ReturnCode(LG_CUSTOM, , False, txtCustomID)
        Case 2                '[3] 품명 코드
            Call ReturnCode(LG_ARTICLE, , False, txtArticle)
''        Case 3                '[4] 오더 코드
''            Call ReturnCode(LG_ORDER, , False, txtOrderNO)
''            txtSearch.Text = txtOrderNO.Tag
''            Call FillStuffOrderData(txtSearch)
''
''        Case 4                '[4] 품명 코드
''            Call ReturnCode(LG_ARTICLE, , False, TxtArticleID2)
    End Select
End Sub

''Private Sub SetClearEdit()
''''    Call ClearData
''''    Call ClearGridSub
''
''    Call ClearScreen(Me, "pnlData")
''
''    cmdFind(1).Enabled = False
''    cmdFind(3).Enabled = False
''    cmdFind(4).Enabled = False
''
''    txtCustomID.Tag = ""
''    TxtArticleID2.Tag = ""
''    txtSearch.Tag = ""
''
''    dtpDate(2) = Now
''    grdData.Rows = grdData.FixedRows
''    grdData.HighLight = flexHighlightNever
''
''    Call SetKeyEdit(True)
''End Sub

Private Sub cmdSearch_Click()
    If optGroup(0) Then
        Call FillGridGroup(True)
    Else
        Call FillGridGroup(False)
    End If
End Sub
Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[3] 금일
        dtpDate(0) = Date - 1
        dtpDate(1) = Date - 1
    ElseIf Index = 1 Then '[2] 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub SetGridGroup(NewFlex As VSFlexGrid)
    With NewFlex
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .GridLines = flexGridNone
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .BackColorBkg = vbWhite
        .SheetBorder = vbWhite
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        .RowHeight(0) = 450
        .RowHeightMin = 275
    End With
End Sub

Private Sub InitGroup(Optional NewValue As Boolean = True)
    Dim i%
    Call SetGridGroup(grdGroup)
    
    For i = 0 To grdGroup.Cols - 1
        grdGroup.ColHidden(i) = False
    Next i
    grdGroup.Redraw = flexRDNone
    
    '----- 오더별 집계 조회
    If NewValue Then
        With grdGroup
            .Redraw = flexRDNone
            
            .Rows = 2
            .FixedRows = 2
            .FixedCols = 0
            .Cols = 18
            .RowHeight(0) = 350
            .RowHeight(1) = 350

            .TextArray(0) = " ":                                .ColWidth(0) = 200
            .TextArray(1) = " ":                                .ColWidth(1) = 200
            .TextArray(2) = "관리번호":                         .ColAlignment(2) = flexAlignCenterCenter
            .TextArray(3) = "Order NO":                         .ColAlignment(3) = flexAlignLeftCenter
            .TextArray(4) = "접수일자":                         .ColAlignment(4) = flexAlignCenterCenter
            .TextArray(5) = "거  래  처":                       .ColAlignment(5) = flexAlignLeftCenter
            .TextArray(6) = "품      명":                       .ColAlignment(6) = flexAlignLeftCenter
            .TextArray(7) = "가공구분":                         .ColAlignment(7) = flexAlignCenterCenter
            .TextArray(8) = "원단폭":                           .ColAlignment(8) = flexAlignCenterCenter
            .TextArray(9) = "축율" & vbCrLf & "LOSS":           .ColAlignment(9) = flexAlignCenterCenter
            .TextArray(10) = "색상수":                          .ColAlignment(10) = flexAlignRightCenter
            .TextArray(11) = "주문량":                          .ColAlignment(11) = flexAlignRightCenter
            .TextArray(12) = "입고" & vbCrLf & "절수":          .ColAlignment(12) = flexAlignRightCenter
            .TextArray(13) = "입고량":                          .ColAlignment(13) = flexAlignRightCenter
            .TextArray(14) = "배색량":                          .ColAlignment(14) = flexAlignRightCenter
            .TextArray(15) = "Sort OrderID":                    .ColAlignment(15) = flexAlignCenterCenter
            .TextArray(16) = "OrderNo":                         .ColAlignment(17) = flexAlignCenterCenter
            .TextArray(17) = "StuffDate":                       .ColAlignment(16) = flexAlignCenterCenter

            .TextArray(.Cols + 0) = " "
            .TextArray(.Cols + 1) = " "
            .TextArray(.Cols + 2) = "관리번호"
            .TextArray(.Cols + 3) = "Order NO"
            .TextArray(.Cols + 4) = "접수일자"
            .TextArray(.Cols + 5) = "입  고  처"
            .TextArray(.Cols + 6) = "입고일자(사종)"
            .TextArray(.Cols + 7) = "가공구분"
            .TextArray(.Cols + 8) = "원단폭"
            .TextArray(.Cols + 9) = "축율" & vbCrLf & "LOSS"
            .TextArray(.Cols + 10) = "색상수"
            .TextArray(.Cols + 11) = "주문량"
            .TextArray(.Cols + 12) = "입고" & vbCrLf & "절수"
            .TextArray(.Cols + 13) = "입고량"
            .TextArray(.Cols + 14) = "배색량"
            .TextArray(.Cols + 15) = "Sort OrderID"
            .TextArray(.Cols + 16) = "OrderNo"
            .TextArray(.Cols + 17) = "StuffDate"

            .ColWidth(0) = 200
            .ColWidth(1) = 200
            .ColWidth(2) = 1400
            .ColWidth(3) = 1400
            .ColWidth(4) = 800
            .ColWidth(5) = 1700
            .ColWidth(6) = 3000
            .ColWidth(7) = 1000
            .ColWidth(8) = 800
            .ColWidth(9) = 1300
            .ColWidth(10) = 600
            .ColWidth(11) = 1000
            .ColWidth(12) = 800
            .ColWidth(13) = 1000
            .ColWidth(14) = 2000
            .ColWidth(15) = 0
            .ColWidth(16) = 0
            .ColWidth(17) = 0
            
            .ColHidden(2) = True
            .ColHidden(15) = True
            .ColHidden(16) = True
            .ColHidden(17) = True
            
            For i = 1 To .Cols - 1
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
            Next i
            .ScrollBars = flexScrollBarVertical
            .Redraw = flexRDDirect
        End With

    Else
        With grdGroup
            .Rows = 2
            .FixedRows = 2
            .FixedCols = 0
            .Cols = 19
            .RowHeight(0) = 350
            .RowHeight(1) = 350
            
            .Redraw = flexRDNone
            
    
            .TextArray(0) = " ":                                       .ColWidth(0) = 100
            .TextArray(1) = " ":                                       .ColWidth(1) = 200
            .TextArray(2) = "거래처ID":                               .ColWidth(2) = 600:              .ColAlignment(2) = flexAlignCenterCenter
            .TextArray(3) = "거래처명":                               .ColWidth(3) = 1400:             .ColAlignment(3) = flexAlignCenterCenter
            .TextArray(4) = "관리번호":                               .ColWidth(4) = 1400:             .ColAlignment(4) = flexAlignCenterCenter
            .TextArray(5) = "Order NO":                               .ColWidth(5) = 1400:             .ColAlignment(5) = flexAlignLeftCenter
            .TextArray(6) = "접 수 일":                               .ColWidth(6) = 1200:              .ColAlignment(6) = flexAlignCenterCenter
            .TextArray(7) = "품    명":                               .ColWidth(7) = 3000:             .ColAlignment(7) = flexAlignLeftCenter
            .TextArray(8) = "가공구분":                               .ColWidth(8) = 1000:              .ColAlignment(8) = flexAlignRightCenter
            .TextArray(9) = "원단폭":                                 .ColWidth(9) = 1000:              .ColAlignment(9) = flexAlignCenterCenter
            .TextArray(10) = "축율" & vbCrLf & "LOSS":                .ColWidth(10) = 1000:            .ColAlignment(10) = flexAlignCenterCenter
            .TextArray(11) = "색상수":                                .ColWidth(11) = 800:             .ColAlignment(11) = flexAlignRightCenter
            .TextArray(12) = "주문량":                                .ColWidth(12) = 800:             .ColAlignment(12) = flexAlignRightCenter
            .TextArray(13) = "입고" & vbCrLf & "절수":                .ColWidth(13) = 800:             .ColAlignment(13) = flexAlignRightCenter
            .TextArray(14) = "입고량":                                .ColWidth(14) = 1000:             .ColAlignment(14) = flexAlignRightCenter
            .TextArray(15) = "배색량":                                .ColWidth(15) = 1400:             .ColAlignment(15) = flexAlignRightCenter:
            .TextArray(16) = "CustomID":                              .ColWidth(16) = 0:               .ColAlignment(16) = flexAlignCenterCenter
            .TextArray(17) = "OrderID":                               .ColWidth(17) = 0:               .ColAlignment(17) = flexAlignCenterCenter
            .TextArray(18) = "Stuff-Pkey":                            .ColWidth(18) = 0:               .ColAlignment(18) = flexAlignCenterCenter
    
    
            .TextArray(.Cols + 0) = " "
            .TextArray(.Cols + 1) = " "
            .TextArray(.Cols + 2) = "거래처ID"
            .TextArray(.Cols + 3) = "거래처명"
            .TextArray(.Cols + 4) = "관리번호"
            .TextArray(.Cols + 5) = "Order NO"
            .TextArray(.Cols + 6) = "입 고 일"
            .TextArray(.Cols + 7) = "입 고 처"
            .TextArray(.Cols + 8) = "가공구분"
            .TextArray(.Cols + 9) = "원단폭"
            .TextArray(.Cols + 10) = "축율" & vbCrLf & "LOSS"
            .TextArray(.Cols + 11) = "색상수"
            .TextArray(.Cols + 12) = "주문량"
            .TextArray(.Cols + 13) = "입고" & vbCrLf & "절수"
            .TextArray(.Cols + 14) = "입고량"
            .TextArray(.Cols + 15) = "배색량"
            .TextArray(.Cols + 16) = "CustomID"
            .TextArray(.Cols + 17) = "OrderID"
            .TextArray(.Cols + 18) = "Stuff-Pkey"
    
            For i = 1 To .Cols - 1
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
            Next i
    
            .ColHidden(2) = True
            .ColHidden(4) = True
            .ColHidden(16) = True
            .ColHidden(17) = True
            .ColHidden(18) = True
            .ScrollBars = flexScrollBarVertical
    
            .Redraw = flexRDDirect
        End With
    End If
    
    With grdGroup
        .MergeCells = flexMergeFixedOnly
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next i
        .Redraw = flexRDDirect
    End With
    
    Call SetToggle
    
    With grdTotal
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByColumn
        .FixedRows = 0
        .Rows = 1
        .Cols = 9
        .ExtendLastCol = True
        
        .RowHeight(0) = 275
        
        .TextArray(0) = "합          계":  .ColWidth(0) = 5000: .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "주문량:":         .ColWidth(1) = 1250: .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "0 YDS":           .ColWidth(2) = 1250: .ColAlignment(2) = flexAlignRightCenter
        
        .TextArray(3) = "입고절수:":       .ColWidth(3) = 1250: .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "0 절":            .ColWidth(4) = 1250: .ColAlignment(4) = flexAlignRightCenter
        
        .TextArray(5) = "입고수량:":       .ColWidth(5) = 1250: .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "0 YDS":           .ColWidth(6) = 1250: .ColAlignment(6) = flexAlignRightCenter
        
        .TextArray(7) = "배색량:":         .ColWidth(7) = 1250: .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "0 YDS":           .ColWidth(8) = 1250: .ColAlignment(8) = flexAlignRightCenter
        
        For i = 1 To 7 Step 2
            .Cell(flexcpForeColor, 0, i, 0, i) = &HFFFFFF
            .Cell(flexcpBackColor, 0, i, 0, i) = &H800000
        Next
         .Redraw = flexRDDirect
    End With
End Sub



''Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, iRow As Integer, iLvl As Integer)
''    With oFlex
''        ' Set the row as a group
''        .IsSubtotal(iRow) = True
''        ' Set the indentation level of the group
''        .RowOutlineLevel(iRow) = iLvl
''
''        Select Case iLvl
''        Case 0
''            .Cell(flexcpForeColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
''            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
''        Case 1, 2
''            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
''        End Select
''    End With
''End Sub



Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub grdGroup_DblClick()
    With grdGroup
        If .Row <= .FixedRows Then Exit Sub

        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If
    End With
End Sub


'*******************************************************************************************
'--- 생지 입고 오더별, 거래처별 조회
'*******************************************************************************************
Private Sub FillGridGroup(Optional NewValue As Boolean = True)
    Dim oStuffIn As PlusLib2.cStuffIN
    Dim rs As ADODB.Recordset
    Dim iTop(2) As Integer, nTop%
    Dim i%, xpName As String
    Dim nCheckNon As Integer
    Dim nTotOrderQty As Long, nTotRoll As Long, nTotQty As Long, nTotColorQty As Long, iSubRow%
    Dim StuffClss As String

'    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.cStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    If NewValue Then
        xpName = "xp_StuffIN_sStuffIN"
    Else
        xpName = "xp_StuffIN_sStuffIN_Custom"
    End If
    
    ' 확정구분
    If chkSearch(5).Value Then
        nCheckNon = cboOrderID.ListIndex + 1
    Else
        nCheckNon = 0  '전체
    End If
    
    If chkSearch(4).Value Then
        StuffClss = CboStuffClss2.ItemData(CboStuffClss2.ListIndex)
    Else
        StuffClss = ""
    End If
    
    iSubRow = 0
    
    Set rs = oStuffIn.GetStuffIN(xpName, IIf(chkSearch(3) = vbChecked, 1, 0) _
                                , MakeDate(DF_SHORT, dtpDate(0)) _
                                , MakeDate(DF_SHORT, dtpDate(1)) _
                                , IIf(chkSearch(1) = vbChecked, 1, 0) _
                                , txtCustom(1).Tag _
                                , IIf(chkSearch(2) = vbChecked, 1, 0) _
                                , txtArticle.Tag _
                                , IIf(chkSearch(4) = vbChecked, 1, 0) _
                                , StuffClss _
                                , IIf(chkSearch(0) = vbChecked, 1, 0) _
                                , txtSearch(3).Text _
                                , nCheckNon, 0)

    Set oStuffIn = Nothing
    
    If rs.RecordCount = 0 Then
        grdGroup.Rows = grdGroup.FixedRows
        Exit Sub
    End If
    
    Call InitGroup(NewValue)
    
    nTotOrderQty = 0: nTotRoll = 0: nTotQty = 0: nTotColorQty = 0
    
    '------- 오더별 집계 조회
    If NewValue Then
        
        With grdGroup
            .Redraw = flexRDNone
            .Rows = .FixedRows
            Do Until rs.EOF
                
                '---- 첫번째 그룹설정 (OrderID)
'                If Trim(rs!OrderID) <> Trim(.TextMatrix(.Rows - 1, 15)) Then

                If Trim(rs!OrderID) & Trim$(rs!Custom1) & Trim$(rs!Article) <> _
                   Trim(.TextMatrix(iSubRow, 15)) & Trim(.TextMatrix(iSubRow, 5)) & Trim(.TextMatrix(iSubRow, 6)) Then
                
                    .AddItem " "
                    iSubRow = .Rows - 1
                    
                    .TextMatrix(.Rows - 1, 2) = IIf(Trim(rs!OrderID) = "*", "", MakeOrderID(rs!OrderID, OM_EXPAND))
                    .TextMatrix(.Rows - 1, 3) = IIf(Trim(rs!OrderNo) = "", "", rs!OrderNo)
                    .TextMatrix(.Rows - 1, 4) = MakeDate(DF_MD, rs!AcptDate)
                    .TextMatrix(.Rows - 1, 5) = Trim(rs!Custom1)
                    .TextMatrix(.Rows - 1, 6) = Trim(rs!Article)
                    .TextMatrix(.Rows - 1, 7) = rs!WorkName
                    .TextMatrix(.Rows - 1, 8) = rs!Width
                    .TextMatrix(.Rows - 1, 9) = MakeRating(rs!ChunkRate, rs!LossRate)
                    .TextMatrix(.Rows - 1, 10) = CheckNum(rs!ColorQty)
                    .TextMatrix(.Rows - 1, 11) = SetCurrency(CheckNum(rs!OrderQty))
                    .TextMatrix(.Rows - 1, 12) = 0
                    .TextMatrix(.Rows - 1, 13) = 0
                    .TextMatrix(.Rows - 1, 14) = rs!배색Qty
                    .TextMatrix(.Rows - 1, 15) = rs!OrderID
                    .TextMatrix(.Rows - 1, 16) = rs!OrderNo
                    .TextMatrix(.Rows - 1, 17) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 1)
'                    call DoFlexGridGroup(grdgroup, .Rows-1,
                    Call GridCollapse(grdGroup, nTop)
                    nTop = .Rows - 1
                    
                    iTop(1) = .Rows - 1
                    
                    nTotOrderQty = nTotOrderQty + CheckNum(rs!OrderQty)
                    nTotColorQty = nTotColorQty + CheckNum(rs!배색Qty)
                End If
'
                .AddItem "" & vbTab & "" & vbTab & "" & vbTab & ""
                .TextMatrix(.Rows - 1, 5) = CheckNull(rs!Custom2)
                .TextMatrix(.Rows - 1, 6) = MakeDate(DF_MID, rs!StuffDate) & "(" + CheckNull(rs!ThreadName) + ")"
                .TextMatrix(.Rows - 1, 12) = rs!StuffRoll
                .TextMatrix(.Rows - 1, 13) = SetCurrency(rs!StuffQty)
                .TextMatrix(.Rows - 1, 15) = rs!OrderID
                .TextMatrix(.Rows - 1, 16) = rs!OrderNo
                .TextMatrix(.Rows - 1, 17) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                         
                         
                '-------입고절수 , 입고수량 Order별로 합계
                .TextMatrix(iTop(1), 12) = SetCurrency(.TextMatrix(iTop(1), 12) + rs!StuffRoll)
                .TextMatrix(iTop(1), 13) = SetCurrency(.TextMatrix(iTop(1), 13) + rs!StuffQty)
                nTotRoll = nTotRoll + CheckNum(rs!StuffRoll)
                nTotQty = nTotQty + CheckNum(rs!StuffQty)
    
                rs.MoveNext
            Loop
            
            .Redraw = flexRDDirect
        End With
        
    Else
        
        With grdGroup
            .Redraw = flexRDNone
            .Rows = .FixedRows
            
            Do Until rs.EOF
            
                '---- 첫번째 그룹설절  CustomID1 확인
                If Trim(rs!customid1) <> Trim(.TextMatrix(.Rows - 1, 16)) Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 2) = rs!customid1
                    .TextMatrix(.Rows - 1, 3) = rs!Custom1
                    .TextMatrix(.Rows - 1, 12) = 0
                    .TextMatrix(.Rows - 1, 13) = 0
                    .TextMatrix(.Rows - 1, 14) = 0
                    .TextMatrix(.Rows - 1, 15) = 0
                    .TextMatrix(.Rows - 1, 16) = rs!customid1
                    .TextMatrix(.Rows - 1, 18) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 1)
                    Call GridCollapse(grdGroup, nTop)
                    nTop = .Rows - 1
                    
                    iTop(1) = .Rows - 1
                End If
                
                '--- 두번째 그룹설절 OrderID 확인
'                If Trim(rs!OrderID) <> Trim(.TextMatrix(.Rows - 1, 17)) Then
                If Trim(rs!OrderID) & Trim(rs!Article) <> Trim(.TextMatrix(.Rows - 1, 17)) & Trim(.TextMatrix(.Rows - 1, 7)) Then
                
                    .AddItem ""
                    
                    .TextMatrix(.Rows - 1, 4) = MakeOrderID(rs!OrderID, OM_EXPAND)
                    .TextMatrix(.Rows - 1, 5) = rs!OrderNo
                    .TextMatrix(.Rows - 1, 6) = MakeDate(DF_MD, rs!AcptDate)
                    .TextMatrix(.Rows - 1, 7) = rs!Article
                    .TextMatrix(.Rows - 1, 8) = rs!WorkName
                    .TextMatrix(.Rows - 1, 9) = rs!Width
                    .TextMatrix(.Rows - 1, 10) = rs!ChunkRate & "+" & rs!LossRate
                    .TextMatrix(.Rows - 1, 11) = rs!ColorQty
                    .TextMatrix(.Rows - 1, 12) = SetCurrency(rs!OrderQty, 0)
                    .TextMatrix(.Rows - 1, 13) = 0
                    .TextMatrix(.Rows - 1, 14) = 0
                    .TextMatrix(.Rows - 1, 15) = rs!배색Qty
                    
                    .TextMatrix(.Rows - 1, 16) = rs!customid1
                    .TextMatrix(.Rows - 1, 17) = rs!OrderID
                    .TextMatrix(.Rows - 1, 18) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 2)
                 '   Call GridCollapse(grdGroup, nTop)
                 '   nTop = .Rows - 1
                    
                    iTop(2) = .Rows - 1
                    nTotOrderQty = nTotOrderQty + rs!OrderQty
                    nTotColorQty = nTotColorQty + rs!배색Qty
                End If
                
                .AddItem ""
                .TextMatrix(.Rows - 1, 6) = MakeDate(DF_MD, rs!StuffDate)
                .TextMatrix(.Rows - 1, 7) = CheckNull(rs!Custom2) & "(" & rs!ThreadName & ")"
                .TextMatrix(.Rows - 1, 12) = 0
                .TextMatrix(.Rows - 1, 13) = rs!StuffRoll
                .TextMatrix(.Rows - 1, 14) = SetCurrency(rs!StuffQty)
                .TextMatrix(.Rows - 1, 15) = 0
                
                .TextMatrix(.Rows - 1, 16) = rs!customid1
                .TextMatrix(.Rows - 1, 17) = rs!OrderID
                .TextMatrix(.Rows - 1, 18) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                
                nTotRoll = nTotRoll + rs!StuffRoll
                nTotQty = nTotQty + rs!StuffQty
                
                For i = 1 To 2
                    .TextMatrix(iTop(i), 12) = SetCurrency(.TextMatrix(iTop(i), 12) + rs!OrderQty)
                    .TextMatrix(iTop(i), 13) = SetCurrency(.TextMatrix(iTop(i), 13) + rs!StuffRoll)
                    .TextMatrix(iTop(i), 14) = SetCurrency(.TextMatrix(iTop(i), 14) + rs!StuffQty)
                    .TextMatrix(iTop(i), 15) = SetCurrency(.TextMatrix(iTop(i), 15) + rs!배색Qty)

                Next i

                rs.MoveNext
            Loop
            
     '       Call ChangeScroll(1)
            
            .Redraw = flexRDDirect
        End With
    End If
    
    If grdGroup.Rows > grdGroup.FixedRows Then
        grdGroup.Row = grdGroup.FixedRows
    Else
        MsgBox LoadResString(203), vbInformation
    End If
    
    rs.Close
    Set rs = Nothing
    
    Call SetToggle
    
    Call GridCollapse(grdGroup, nTop)
    
    
    With grdTotal
        .TextMatrix(0, 2) = Format(nTotOrderQty, "#,##0 YDS")
        .TextMatrix(0, 4) = Format(nTotRoll, "#,##0 절")
        .TextMatrix(0, 6) = Format(nTotQty, "#,##0 YDS")
        .TextMatrix(0, 8) = Format(nTotColorQty, "#,##0 YDS")
        .Redraw = flexRDDirect
    End With
    
    Exit Sub

ErrHandler:
    grdGroup.Redraw = flexRDDirect
    Call ErrorBox(Err.Number, "frmStuffINView.FillGridGroup", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
End Sub



Private Sub optGroup_Click(Index As Integer)
    
    If Index = 0 Then
        m_bGroupClss = True
        Call InitGroup
    Else
        m_bGroupClss = False
        Call InitGroup(m_bGroupClss)
    End If
    Call FillGridGroup(m_bGroupClss)
End Sub

Private Sub optOrder_Click(Index As Integer)
    Call SetToggle
'    Select Case Index
'        Case 2
'            With grdGroup
'                If m_bGroupClss Then
'                    .ColHidden(2) = True
'                    .ColHidden(3) = False
'
'                Else
'                    .ColHidden(4) = True
'                    .ColHidden(5) = False
'                End If
'            End With
'
'        Case 3
'            With grdGroup
'                If m_bGroupClss Then
'                    .ColHidden(2) = False
'                    .ColHidden(3) = True
'                Else
'                    .ColHidden(4) = False
'                    .ColHidden(5) = True
'                End If
'            End With
'    End Select
End Sub

Sub SetToggle()
    Dim Index As Integer
    If optOrder(2).Value Then
        Index = 2
    Else
        Index = 3
    End If
    
    
    With grdGroup
''        .ColHidden(2) = False
''        .ColHidden(3) = False
''        .ColHidden(4) = False
''        .ColHidden(5) = False
        
        Select Case Index
            Case 2
                '오더별
                If m_bGroupClss Then
                    .ColHidden(2) = True
                    .ColHidden(3) = False
                Else
                    .ColHidden(4) = True
                    .ColHidden(5) = False
                End If
            Case 3
                If m_bGroupClss Then
                    .ColHidden(2) = False
                    .ColHidden(3) = True
                Else
                    .ColHidden(4) = False
                    .ColHidden(5) = True
                End If
        End Select
    End With
End Sub

''Private Sub tabform_Click(PreviousTab As Integer)
''    If PreviousTab = 1 Then
''        Call ChangeMode(Me, True)
''        pnlData.Enabled = False
''    End If
''
''    Select Case tabForm.Tab
''        Case 0
''            Call SetClearEdit
''''          Call cmdSearch_Click
''''        Case 1
''''            If Trim(pnlMsg) = "" Or val(txtStuffSeq.Text) = 0 Then
''''                Call cmdOperate_Click(0)
''''            End If
''    End Select
''
''End Sub

Private Sub txtArticle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_ARTICLE, , False, txtArticle)
        Call MoveFocus(KeyAscii)
    End If

End Sub



Private Sub txtCustom_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            Call MoveFocus(KeyAscii)
        Case 1
            If KeyAscii = vbKeyReturn Then
                Call ReturnCode(LG_CUSTOM, , False, txtCustom(Index))
                Call MoveFocus(KeyAscii)
            End If
    End Select
End Sub



