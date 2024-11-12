VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutwareOn 
   ClientHeight    =   9255
   ClientLeft      =   1665
   ClientTop       =   1530
   ClientWidth     =   11850
   Icon            =   "frmOutwareOn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   420
      TabIndex        =   31
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
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   120
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "전체 선택"
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   30
      Top             =   8490
      Width           =   1140
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "선택 해제"
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   29
      Top             =   8865
      Width           =   1140
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   690
      Left            =   8340
      TabIndex        =   26
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      확인(&S)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10110
      TabIndex        =   25
      Top             =   8460
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
      Top             =   960
      Width           =   11835
      _cx             =   20876
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
   Begin Threed.SSFrame frmSearch 
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1667
      _Version        =   196609
      Begin VB.ComboBox cboBoOutClss 
         Height          =   300
         Left            =   9630
         Style           =   2  '드롭다운 목록
         TabIndex        =   28
         Top             =   465
         Width           =   1185
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
         Left            =   9630
         TabIndex        =   5
         Top             =   75
         Width           =   1185
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   6480
         TabIndex        =   4
         Top             =   465
         Width           =   1365
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   6480
         TabIndex        =   3
         Top             =   75
         Width           =   1365
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
         Left            =   3450
         TabIndex        =   7
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23789569
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   3435
         TabIndex        =   8
         Top             =   465
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23789569
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
            Caption         =   "작성 일자"
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
         Width           =   1260
         _ExtentX        =   2223
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
         Top             =   465
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
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   8310
         TabIndex        =   17
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
            Caption         =   "관리번호"
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1185
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
         TabIndex        =   27
         Top             =   465
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
         Caption         =   "보관분 구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
         Top             =   525
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmOutwareOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_nSelected As Integer
Private m_bloading As Boolean

Private Sub cboBoOutClss_Click()
    If m_bloading Then Exit Sub
    
    If cboBoOutClss.ListIndex = 1 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
    Call FillGridData
End Sub

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

Private Sub cmdCheck_Click(Index As Integer)
    Dim SetValue, i%
    
    If Index = 0 Then   '[0] 전체선택
        SetValue = flexChecked
        cmdSave.Enabled = True
    Else                '[1] 선택 해제
        SetValue = flexUnchecked
        cmdSave.Enabled = False
    End If

    With grdData
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, i, 1) = SetValue
        Next i
    End With
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

Private Sub cmdSave_Click()
    Dim tOw() As PlusLib2.TOUTWARE
    Dim oOutware As PlusLib2.COutWare
    Dim sMsg$
    Dim nPoint%, i%

    Dim oOrder As PlusLib2.COrder
    
    On Error GoTo ErrHandler
    
    If m_nSelected < 1 Then Exit Sub
    
    sMsg = "선택한 보관분에 대해서 보관분으로 확정하시겠습니까?"
    
    If MsgBox(sMsg, vbQuestion + vbYesNo) = vbYes Then
        ReDim tOw(m_nSelected - 1)
        With grdData
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 1) = flexChecked Then
                    tOw(nPoint).OrderID = MakeOrderID(.TextMatrix(i, 4), OM_REDUCE)
                    tOw(nPoint).OutSeq = .TextMatrix(i, 14)
                    
                    nPoint = nPoint + 1
                End If
            Next i
        End With
        
        Set oOutware = New PlusLib2.COutWare
        oOutware.Connection = g_adoCon
        oOutware.UserName = g_sUserName
        
        If oOutware.UpdateBoOutClss(tOw()) Then
            MsgBox "보관분으로 확정하였습니다.", vbInformation
        Else
            MsgBox "보관분 확정에 실패하였습니다", vbCritical
        End If
        Set oOutware = Nothing
        
        Call FillGridData
    End If
    Exit Sub
ErrHandler:
    Set oOrder = Nothing

    Call ErrorBox(Err.Number, "frmOutwareOn.cmdSave_Click", Err.Description)
End Sub

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
    
    m_bloading = True
    With cboBoOutClss
        .AddItem "1. 전체"
        .AddItem "2. 보관분"
        .AddItem "3. 출고분"
        
        .ListIndex = 0
    End With
    m_bloading = False
    
    pnlProgress.Visible = False
    cmdSave.Picture = LoadResPicture("CHECK", vbResIcon)
    cmdSave.Enabled = False
End Sub


Private Sub grdData_KeyPress(KeyAscii As Integer)
    If cboBoOutClss.ListIndex <> 1 Then Exit Sub
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        With grdData
            If .Rows = .FixedRows Or .TextMatrix(.Row, 15) = "*" Then Exit Sub
        End With
        
        Call CheckCount
    End If
End Sub

Private Sub grdData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdData
        If .Rows = .FixedRows Or .MouseRow < .FixedRows Or .MouseRow >= .Rows Or .MouseCol <> 1 Then Exit Sub
        If cboBoOutClss.ListIndex <> 1 Then Exit Sub
        If .TextMatrix(.Row, 15) = "*" Then Exit Sub
    End With

    Call CheckCount
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdData
        If optOrder(0).Value Then
            .ColWidth(5) = 1350
            .ColWidth(4) = 0
            chkSearch(3).Caption = "Order No."
        Else
            .ColWidth(5) = 0
            .ColWidth(4) = 1350
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
    With grdData
        .Cols = 16
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1

        .TextArray(0) = " "
        .TextArray(1) = " ":            .ColWidth(1) = 250:         .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "거래처":       .ColWidth(2) = 1300:        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "품명":         .ColWidth(3) = 1300:        .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "관리번호":     .ColWidth(4) = 1350:        .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "Order No.":    .ColWidth(5) = 0:           .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "가공구분":     .ColWidth(6) = 1000:        .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "수주량":       .ColWidth(7) = 800:         .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "작성일자":     .ColWidth(8) = 1000:        .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "출고절수":     .ColWidth(9) = 850:         .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "출고수량":    .ColWidth(10) = 850:        .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "출고처":      .ColWidth(11) = 1000:       .ColAlignment(11) = flexAlignLeftCenter
        .TextArray(12) = "출고" & vbCrLf & "단가":    .ColWidth(12) = 650:        .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "구분":        .ColWidth(13) = 800:        .ColAlignment(13) = flexAlignCenterCenter
        .TextArray(14) = "OutSeq":      .ColWidth(14) = 0
        .TextArray(15) = "보관분확정":  .ColWidth(15) = 0
        
        .ColFormat(7) = "#,##0"
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        
        .RowHeightMin = 390
        
        .ColDataType(1) = flexDTBoolean
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub FillGridData()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim i%, sBoOutClss$
    
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Set rs = oOutware.GetOutwareOn(IIf(chkSearch(0) = vbChecked, 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                                 IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3), cboBoOutClss.ListIndex)
    Set oOutware = Nothing
        
    With grdData
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            If rs!BoOutClss = "*" Then
                If rs!BoConfirmClss = "*" Then
                    sBoOutClss = "보관분(O)"
                Else
                    sBoOutClss = "보관분(X)"
                End If
            Else
                sBoOutClss = "출고분"
            End If
            .AddItem CStr(.Rows) & vbTab & False & vbTab & rs!KCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!OrderNo & vbTab & rs!WorkName & vbTab & rs!OrderQty & vbTab & MakeDate(DF_LONG, rs!OutDate) & vbTab & rs!OutRoll & vbTab & _
                rs!OutQty & vbTab & rs!OutCustom & vbTab & rs!UnitPrice & vbTab & sBoOutClss & vbTab & rs!OutSeq & vbTab & rs!BoConfirmClss
            
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
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    
    Exit Sub

ErrHandler:
    pnlProgress.Visible = False
    Set oOutware = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmOutwareOn.FillGridData", Err.Description)
End Sub

Private Sub CheckCount()
    With grdData
        If .Cell(flexcpChecked, .Row, 1) = flexUnchecked Then
            .Cell(flexcpChecked, .Row, 1) = flexChecked
            m_nSelected = m_nSelected + 1
        Else
            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
            m_nSelected = m_nSelected - 1
        End If
    End With
    
    If cboBoOutClss.ListIndex = 1 Then
        cmdSave.Enabled = IIf(m_nSelected > 0, True, False)
    End If
End Sub



