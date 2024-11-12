VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmControlStock 
   Caption         =   "재고 입력"
   ClientHeight    =   9255
   ClientLeft      =   3450
   ClientTop       =   3090
   ClientWidth     =   11850
   Icon            =   "frmControlStock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin VB.CommandButton cmdOperate 
      Caption         =   "취소(&C)"
      Height          =   720
      Index           =   4
      Left            =   8655
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   31
      ToolTipText     =   "자료 취소"
      Top             =   30
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "수정(&U)"
      Height          =   720
      Index           =   1
      Left            =   10245
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   30
      ToolTipText     =   "자료 수정"
      Top             =   30
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "삭제(&D)"
      Height          =   720
      Index           =   2
      Left            =   11040
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   29
      ToolTipText     =   "자료 삭제"
      Top             =   30
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "추가(&A)"
      Height          =   720
      Index           =   0
      Left            =   9450
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   0
      ToolTipText     =   "자료 추가"
      Top             =   30
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "저장(&S)"
      Height          =   720
      Index           =   3
      Left            =   7860
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   11
      ToolTipText     =   "자료 저장"
      Top             =   30
      Visible         =   0   'False
      Width           =   780
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   780
      Left            =   30
      TabIndex        =   12
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   1376
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   60
         Width           =   630
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   390
         Width           =   630
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   3810
         TabIndex        =   18
         Top             =   390
         Width           =   2055
      End
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   3810
         TabIndex        =   17
         Top             =   105
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   720
         TabIndex        =   13
         Top             =   90
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   72417281
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   390
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   72417281
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   2580
         TabIndex        =   19
         Top             =   75
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거래처"
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   75
            Width           =   885
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   0
         Left            =   5880
         TabIndex        =   21
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
         Index           =   3
         Left            =   2580
         TabIndex        =   22
         Top             =   390
         Width           =   1200
         _ExtentX        =   2117
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
            Caption         =   "품   명"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   885
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   5880
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   390
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
      Begin Threed.SSCommand cmdSearch 
         Height          =   600
         Left            =   6270
         TabIndex        =   25
         Top             =   90
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1058
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
         Caption         =   "        검색(&F)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   2025
         TabIndex        =   16
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   2025
         TabIndex        =   15
         Top             =   465
         Width           =   360
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdStockList 
      Height          =   6465
      Left            =   30
      TabIndex        =   28
      Top             =   780
      Width           =   11790
      _cx             =   20796
      _cy             =   11404
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
   Begin Threed.SSPanel pnlEdit 
      Height          =   1215
      Left            =   30
      TabIndex        =   32
      Top             =   7260
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   2143
      _Version        =   196609
      Enabled         =   0   'False
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboSubulWidth 
         Height          =   300
         Left            =   4770
         Style           =   2  '드롭다운 목록
         TabIndex        =   46
         Top             =   810
         Width           =   1575
      End
      Begin VB.ComboBox cboUnitClss 
         Height          =   300
         Left            =   8280
         Style           =   2  '드롭다운 목록
         TabIndex        =   10
         Top             =   465
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.ComboBox cboProcInClss 
         Height          =   300
         ItemData        =   "frmControlStock.frx":000C
         Left            =   1170
         List            =   "frmControlStock.frx":000E
         Style           =   2  '드롭다운 목록
         TabIndex        =   7
         Top             =   1245
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cboStockClss 
         Height          =   300
         Left            =   1170
         Style           =   2  '드롭다운 목록
         TabIndex        =   8
         Top             =   1575
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame fraDate 
         Caption         =   "기준일"
         Height          =   675
         Left            =   120
         TabIndex        =   33
         Top             =   60
         Width           =   1815
         Begin VB.OptionButton optDate 
            Caption         =   "현재일"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   35
            Top             =   330
            Width           =   885
         End
         Begin VB.OptionButton optDate 
            Caption         =   "말일"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   330
            Value           =   -1  'True
            Width           =   705
         End
      End
      Begin MRPPlus2.WizText txtCustom 
         Height          =   300
         Left            =   4770
         TabIndex        =   3
         Top             =   105
         Width           =   1815
         _ExtentX        =   3201
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
         MaxLength       =   35
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   36
         Top             =   870
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "기 준 일"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   3540
         TabIndex        =   37
         Top             =   105
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "거 래 처"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   3540
         TabIndex        =   38
         Top             =   465
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "품     명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtArticle 
         Height          =   300
         Left            =   4770
         TabIndex        =   5
         Top             =   465
         Width           =   1815
         _ExtentX        =   3201
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
         MaxLength       =   20
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   5
         Left            =   90
         TabIndex        =   39
         Top             =   1245
         Visible         =   0   'False
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "입고가공"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   6
         Left            =   7050
         TabIndex        =   40
         Top             =   120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "재 고 량"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   7
         Left            =   90
         TabIndex        =   41
         Top             =   1575
         Visible         =   0   'False
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "재고구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtStockQty 
         Height          =   300
         Left            =   8280
         TabIndex        =   9
         Top             =   105
         Width           =   1245
         _ExtentX        =   2196
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
         MaxLength       =   20
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   8
         Left            =   7050
         TabIndex        =   42
         Top             =   465
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "단      위"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpBaseDate 
         Height          =   300
         Left            =   1980
         TabIndex        =   1
         Top             =   105
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Format          =   70844417
         CurrentDate     =   36871
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   6600
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   105
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   3
         Left            =   6600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   465
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   300
         Left            =   1980
         TabIndex        =   43
         Top             =   435
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkTaxClss 
            Caption         =   "사용구분"
            Height          =   195
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   1095
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   9
         Left            =   3540
         TabIndex        =   45
         Top             =   810
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "원 단 폭"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   44
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmControlStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_sFlag        As String * 1

Private Sub cboProcInClss_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        Call NextFocus
    End If
End Sub

Private Sub cboStockClss_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        Call NextFocus
    End If
End Sub

Private Sub SetStuffWidth()
    Dim oCode As PlusLib2.CCode
    Dim rs    As ADODB.Recordset
    Dim II%
    
    On Error GoTo ErrHandler

    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon

    Set rs = oCode.GetStuffWidth
    Set oCode = Nothing
    II = 0
    
    cboSubulWidth.Clear
    If Not rs Is Nothing Then
        If Not rs.BOF Then
           rs.MoveFirst
           Do Until rs.EOF
            cboSubulWidth.AddItem Trim$(rs(0))
            cboSubulWidth.ItemData(II) = val(rs(1))
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

    Err.Raise Err.Number, "Start.MakeCodeCombo", Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub cboUnitClss_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        Call NextFocus
    End If
End Sub

Private Sub chkSearch_Click(Index As Integer)

    If chkSearch(Index).Value = vbChecked Then
        txtSearch(Index).Enabled = True
        txtSearch(Index).SetFocus
        cmdFind(Index).Enabled = True
    Else
        txtSearch(Index).Enabled = False
        cmdFind(Index).Enabled = False
    End If
    
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 0 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(0))
    ElseIf Index = 1 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(1))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtCustom)
    ElseIf Index = 3 Then
        Call ReturnCode(LG_ARTICLE, , False, txtArticle)
    End If

End Sub

Private Sub cmdOperate_Click(Index As Integer)

    On Error GoTo ErrHandler
    
    Select Case Index
        Case ID_ADDNEW
            m_sFlag = ID_ADDNEW
            Call ClearData
            Call ChangeMode(Me, False)
            
            pnlEdit.Enabled = True
            fraDate.Visible = True
            cmdFind(2).Enabled = True
            cmdFind(3).Enabled = True
            chkTaxClss.Enabled = True
            cboSubulWidth.Enabled = True
            
            dtpBaseDate.SetFocus
            
        Case ID_UPDATE
'            If (grdStockList.Rows = grdStockList.FixedRows) Or (grdStockList.Row < grdStockList.FixedRows) Then
'                MsgBox "해당 건을 선택한 후 버튼을 눌러주세요", vbInformation, "항목 선택"
'                Exit Sub
'            End If
            
            m_sFlag = ID_UPDATE
            Call ChangeMode(Me, False)
            pnlEdit.Enabled = True
            fraDate.Visible = False
            dtpBaseDate.Enabled = False
            txtCustom.Enabled = False
            txtArticle.Enabled = False
            chkTaxClss.Enabled = False
            cboSubulWidth.Enabled = False
'            cboProcInClss.SetFocus

        Case ID_DELETE
'            If (grdStockList.Rows = grdStockList.FixedRows) Or (grdStockList.Row < grdStockList.FixedRows) Then
'                MsgBox "해당 건을 선택한 후 버튼을 눌러주세요", vbInformation, "항목 선택"
'                Exit Sub
'            End If
            fraDate.Visible = False
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                m_sFlag = ID_DELETE
                If SaveData Then
                    Call ClearData
                    Call FillGrid
                    m_sFlag = ""
                End If
            End If
        Case ID_SAVE
            If Not CheckData() Then Exit Sub
    
            If SaveData() Then
                Call ChangeMode(Me, True)
                Call ClearData
                pnlEdit.Enabled = False
                fraDate.Visible = False
                cmdFind(2).Enabled = False
                cmdFind(3).Enabled = False
                Call FillGrid
                m_sFlag = ""
            End If
        Case ID_CANCEL
            m_sFlag = ""
            Call ChangeMode(Me, True)
            Call ClearData
            pnlEdit.Enabled = False
            fraDate.Visible = False
            cmdFind(2).Enabled = False
            cmdFind(3).Enabled = False
            Call grdStockList_RowColChange
        End Select

    Exit Sub
ErrHandler:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Err.Clear

End Sub

Private Function SaveData() As Boolean
    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim TSbStock As PlusLib2.TSbStock
    
    On Error GoTo ErrHandler
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    oSubul.UserName = g_sUserName
    
    With TSbStock
        .sBasisDate = MakeDate(DF_SHORT, dtpBaseDate)
        .sCustomID = txtCustom.Tag
        .sArticleID = txtArticle.Tag
        .sSubulWdithID = Format(cboSubulWidth.ItemData(cboSubulWidth.ListIndex), "00") '가공 폭
        .sTaxClss = IIf(chkTaxClss.Value = vbChecked, "1", "0")
        .nStockQty = CheckNum(txtStockQty)
        .sStockUnitClss = "1"
    End With
    
    
    Select Case m_sFlag
        Case ID_ADDNEW
            Set rs = oSubul.GetStockDataOne(TSbStock)
            
            If Not rs.EOF Then
                MsgBox "해당 기준일, 거래처, 품명의 데이터가 이미 존재합니다" & vbCrLf & vbCrLf & _
                        "확인후 다시 작업해주시기 바랍니다", vbCritical, "기 등록건"
                rs.Close
                Set rs = Nothing
                SaveData = True
            Else
                If oSubul.AddNewStock(TSbStock) Then
                    SaveData = True
                    MsgBox "정상적으로 입력되었습니다", vbInformation + vbOKOnly, "입력 성공"
                End If
            End If
            
        Case ID_UPDATE
            If oSubul.UpdateStock(TSbStock) Then
                SaveData = True
                MsgBox "정상적으로 수정되었습니다", vbInformation + vbOKOnly, "수정 성공"
            End If
        Case ID_DELETE
            If oSubul.DeleteStock(TSbStock) Then
                SaveData = True
                MsgBox "정상적으로 삭제되었습니다", vbInformation + vbOKOnly, "삭제 성공"
            End If
            
    End Select
    
    Set oSubul = Nothing
    Exit Function
ErrHandler:
    Set oSubul = Nothing

    Call ErrorBox(Err.Number, "frmControlStock.SaveData", Err.Description)
End Function

Private Sub cmdSearch_Click()
    Call FillGrid
End Sub

Private Sub dtpBaseDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub


Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11975, 9660
    
    Call SetOperate(Me)
    For i = 0 To 3
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        cmdFind(i).Enabled = False
    Next i
    CboListSet
    Call InitGrid
    Call SetStuffWidth
'    Call ClearData
    
End Sub

Private Sub ClearData()
    dtpBaseDate.Enabled = True
    txtCustom.Enabled = True
    txtCustom.Text = ""
    txtCustom.Tag = ""
    txtArticle.Enabled = True
    txtArticle.Text = ""
    txtArticle.Tag = ""
    txtStockQty.Text = ""
    cboProcInClss.ListIndex = 0
    cboStockClss.ListIndex = 0
    cboUnitClss.ListIndex = 0
    chkTaxClss.Value = vbUnchecked
End Sub
Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CboListSet()
Dim oCode As PlusLib2.CCode
Dim rs As ADODB.Recordset
Dim i%

    For i = 0 To 1
        dtpDate(i) = Now
    Next i

    dtpBaseDate = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    
    With cboStockClss
        .Clear
        .AddItem "생지"
        .ItemData(.NewIndex) = "1"
        .AddItem "재고"
        .ItemData(.NewIndex) = "2"
        .ListIndex = -1
    End With
        
    With cboUnitClss
        .Clear
        .AddItem "YDS"
        .ItemData(.NewIndex) = "1"
        .AddItem "MTS"
        .ItemData(.NewIndex) = "2"
        .AddItem "KGS"
        .ItemData(.NewIndex) = "3"
        .ListIndex = -1
    End With
    
    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon
    
    oCode.CodeType = CD_WORK
    
    Set rs = oCode.GetCode()
    Set oCode = Nothing

    With cboProcInClss
        .Clear
        
        For i = 1 To rs.RecordCount
            .AddItem rs!WorkName
            .ItemData(.NewIndex) = rs!WorkID
            
            rs.MoveNext
        Next i
        
        rs.Close
        Set rs = Nothing
        .ListIndex = -1
    End With
    

    
    
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[1] 금일
        dtpDate(0) = Date
        dtpDate(1) = Date
    ElseIf Index = 1 Then '[2] 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If
End Sub

Private Function CheckData() As Boolean
    Dim i%
    CheckData = True
    If m_sFlag = ID_ADDNEW Then
    
    End If
    
    If Trim(txtCustom) = "" Then
        MsgBox "거래처를 입력해야 합니다", vbInformation, "거래처 입력 요망"
        CheckData = False
    End If
    If Trim(txtArticle) = "" Then
        MsgBox "품명을 입력해야 합니다", vbInformation, "품명 입력 요망"
        CheckData = False
    End If
'    If cboProcInClss.ListIndex < 0 Then
'        MsgBox "입고가공 구분을 선택해야 합니다", vbInformation, "입고가공 선택 요망"
'        CheckData = False
'    End If
'    If cboStockClss.ListIndex < 0 Then
'        MsgBox "재고 구분을 선택해야 합니다", vbInformation, "재고구분 선택 요망"
'        CheckData = False
'    End If
    If Trim(txtStockQty) = "" Then
        MsgBox "재고량을 입력해야 합니다", vbInformation, "재고량 입력 요망"
        CheckData = False
    End If
    If Not IsNumeric(txtStockQty) Then
        MsgBox "재고량은 숫자데이터이어야 합니다", vbInformation, "숫자 입력 요망"
        CheckData = False
    End If
        
'    If cboUnitClss.ListIndex < 0 Then
'        MsgBox "단위를 선택해야 합니다", vbInformation, "단위 선택 요망"
'        CheckData = False
'    End If
        
End Function

Private Sub InitGrid()
    Dim iCol%
    
'    Call SetVSFlexGrid(grdStockList)
    With grdStockList
        .Redraw = flexRDNone
        
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .WordWrap = False
        .ExtendLastCol = True
        
        .Cols = 27:     .Rows = 3
        .FixedCols = 1: .FixedRows = 3
        
        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .RowHeight(2) = 300

        For iCol = 0 To .Cols - 1
            .ColWidth(iCol) = 0
            .FixedAlignment(iCol) = flexAlignCenterCenter
        Next iCol
        
        .TextMatrix(2, 0) = "":                     .ColWidth(0) = 300:         .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(2, 1) = "기준일":               .ColWidth(1) = 1400:        .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(2, 2) = "거래처":               .ColWidth(2) = 2800:        .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(2, 3) = "품명":                 .ColWidth(3) = 3000:        .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(2, 4) = "원단폭":               .ColWidth(4) = 1000:        .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(2, 5) = "재고구분":             .ColWidth(5) = 0:           .ColAlignment(5) = flexAlignCenterCenter  '
        .TextMatrix(2, 6) = "재  고  량":           .ColWidth(6) = 2000:        .ColAlignment(6) = flexAlignRightCenter   '재고량
        .TextMatrix(2, 7) = "재  고  량":           .ColWidth(7) = 700:         .ColAlignment(7) = flexAlignCenterCenter   '단위
        .TextMatrix(2, 8) = "재  고  량":           .ColWidth(8) = 400:         .ColAlignment(8) = flexAlignCenterCenter   '사용구분 a/ b
        
        .TextMatrix(2, 21) = "BasisDate"
        .TextMatrix(2, 22) = "CustomID"
        .TextMatrix(2, 23) = "ArticleID"
        .TextMatrix(2, 24) = "SubulWidthID"
        .TextMatrix(2, 25) = "TaxClss"
        .TextMatrix(2, 26) = "StockUnitClss"
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(2) = True
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FillGrid()
    Dim oSubul As PlusLib2.CSubul
    Dim rs As Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    Set rs = oSubul.GetStockList(MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                            IIf(chkSearch(0) = vbChecked, 1, 0), txtSearch(0).Tag, IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag)
    
    Set oSubul = Nothing
        
    With grdStockList
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 1 To rs.RecordCount
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            
            .TextMatrix(.Rows - 1, 0) = CStr(i)
            .TextMatrix(.Rows - 1, 1) = MakeDate(DF_MD, rs!BasisDate)
            .TextMatrix(.Rows - 1, 2) = Trim(rs!kCustom)
            .TextMatrix(.Rows - 1, 3) = Trim(rs!Article)
            .TextMatrix(.Rows - 1, 4) = Trim(rs!StuffWidth)
'            .TextMatrix(.Rows - 1, 5) = IIf(rs!StockClss = "1", "생지", "재고")
            .TextMatrix(.Rows - 1, 6) = Format(rs!StockQty, "##,##0")
            .TextMatrix(.Rows - 1, 7) = IIf(rs!StockUnitClss = "1", "YDS", IIf(rs!StockUnitClss = "2", "MTS", "KGS"))
            .TextMatrix(.Rows - 1, 8) = IIf(rs!TaxClss = "1", "■", "")
            
            .TextMatrix(.Rows - 1, 21) = rs!BasisDate
            .TextMatrix(.Rows - 1, 22) = rs!CustomID
            .TextMatrix(.Rows - 1, 23) = rs!ArticleID
            .TextMatrix(.Rows - 1, 24) = rs!SubulWidthID
'            .TextMatrix(.Rows - 1, 25) = rs!T
'            .TextMatrix(.Rows - 1, 26) = rs!StockUnitClss
            
            rs.MoveNext
        Next i
        rs.Close
        
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If
        .Select 1, 1, 1, 2
'        .Sort = flexSortGenericAscending
        .Redraw = flexRDDirect
    End With
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmControl.FillGrid", Err.Description)
End Sub


Private Sub grdStockList_RowColChange()
    Dim sDate$

    With grdStockList
        If .Rows > .FixedRows And .Row >= .FixedRows Then
        
            txtCustom.Text = .TextMatrix(.Row, 2)
            txtCustom.Tag = .TextMatrix(.Row, 22)
            txtArticle.Text = .TextMatrix(.Row, 3)
            txtArticle.Tag = .TextMatrix(.Row, 23)
            txtStockQty.Text = .TextMatrix(.Row, 6)
            chkTaxClss.Value = IIf(.TextMatrix(.Row, 8) = "", vbUnchecked, vbChecked)
            sDate = .TextMatrix(.Row, 21)
'            sDate = Left(sDate, 4) & "-" & Mid(sDate, 5, 2) & "-" & Right(sDate, 2)
            
            cboSubulWidth.ListIndex = FindComboBox(cboSubulWidth, .TextMatrix(.Row, 24))
            dtpBaseDate.Value = CDate(Format(sDate, "0000/00/00"))
'            cboProcInClss.ListIndex = FindComboBox(cboProcInClss, CLng("0" & .TextMatrix(.Row, 24)))
'            cboStockClss.ListIndex = FindComboBox(cboStockClss, CLng("0" & .TextMatrix(.Row, 25)))
'            cboUnitClss.ListIndex = FindComboBox(cboUnitClss, CLng("0" & .TextMatrix(.Row, 26)))
        End If
    
    End With
End Sub

Private Sub optDate_Click(Index As Integer)
    If Index = 0 Then
        dtpBaseDate = Now
    Else
        dtpBaseDate = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If
End Sub

Private Sub txtArticle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then

        Call ReturnRef(LG_ARTICLE, , False, txtArticle)
        
'        If Len(txtArticle.Tag) = 0 Then
'            cboProcInClss.SetFocus
'        End If
    End If
End Sub

Private Sub txtCustom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnRef(LG_CUSTOM, , False, txtCustom)
        
'        If Len(txtCustom.Tag) = 0 Then
            txtArticle.SetFocus
'        End If
    End If

End Sub

            
Private Sub txtStockQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(txtStockQty.Text) Then
            MsgBox "숫자가 잘못되어 있습니다", vbCritical + vbOKOnly, "숫자입력 에러"
            txtStockQty.SetFocus
            Exit Sub
        End If
        txtStockQty.Text = Format(txtStockQty.Text, "##,##0")
'        cboUnitClss.SetFocus
    End If

End Sub
