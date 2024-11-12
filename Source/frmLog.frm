VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLog 
   ClientHeight    =   9255
   ClientLeft      =   1410
   ClientTop       =   1590
   ClientWidth     =   11865
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전월"
      Height          =   315
      Index           =   0
      Left            =   1485
      MousePointer    =   99  '사용자 정의
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   630
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   2130
      MousePointer    =   99  '사용자 정의
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   630
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   2
      Left            =   1485
      MousePointer    =   99  '사용자 정의
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   450
      Width           =   630
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금년"
      Height          =   315
      Index           =   3
      Left            =   2130
      MousePointer    =   99  '사용자 정의
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   450
      Width           =   630
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   765
      Left            =   10965
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   14
      ToolTipText     =   "자료 저장"
      Top             =   75
      Width           =   780
   End
   Begin VB.ComboBox cboSearch 
      Height          =   300
      Index           =   1
      Left            =   7305
      Style           =   2  '드롭다운 목록
      TabIndex        =   12
      Top             =   135
      Width           =   1710
   End
   Begin VB.ComboBox cboSearch 
      Height          =   300
      Index           =   2
      Left            =   7305
      Style           =   2  '드롭다운 목록
      TabIndex        =   11
      Top             =   480
      Width           =   1710
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "전체 선택"
      Height          =   315
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   1140
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "선택 해제"
      Height          =   315
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   480
      Width           =   1140
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   360
      TabIndex        =   2
      Top             =   3420
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
         TabIndex        =   3
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
         TabIndex        =   4
         Top             =   120
         Width           =   270
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   270
      Index           =   0
      Left            =   5340
      TabIndex        =   5
      Top             =   150
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   476
      _Version        =   196609
      Caption         =   "부터"
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   300
      Index           =   1
      Left            =   5340
      TabIndex        =   6
      Top             =   480
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   529
      _Version        =   196609
      Caption         =   "까지"
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   6030
      TabIndex        =   7
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "사용자"
         Height          =   240
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   0
      Left            =   6030
      TabIndex        =   9
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "컴퓨터"
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   10
         Top             =   45
         Width           =   1080
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7515
      Left            =   30
      TabIndex        =   13
      Top             =   855
      Width           =   11790
      _cx             =   20796
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
      MousePointer    =   1
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   4065
      TabIndex        =   19
      Top             =   120
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23789569
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   4065
      TabIndex        =   20
      Top             =   465
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      Format          =   23789569
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   2820
      TabIndex        =   21
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "로그일자"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   690
      Left            =   8340
      TabIndex        =   23
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      삭제(&D)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10155
      TabIndex        =   24
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   8850
      Left            =   15
      TabIndex        =   25
      Top             =   15
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   15610
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   750
      TabCaption(0)   =   "   일반 로그   "
      TabPicture(0)   =   "frmLog.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line2(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "   오류 로그   "
      TabPicture(1)   =   "frmLog.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2(0)"
      Tab(1).Control(1)=   "Line2(1)"
      Tab(1).ControlCount=   2
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         Index           =   3
         X1              =   1335
         X2              =   1335
         Y1              =   0
         Y2              =   840
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         Index           =   2
         X1              =   1350
         X2              =   1350
         Y1              =   -15
         Y2              =   825
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   -73665
         X2              =   -73665
         Y1              =   15
         Y2              =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   -73650
         X2              =   -73650
         Y1              =   0
         Y2              =   840
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bSortForward As Boolean
Private m_nSelected    As Integer

Private Sub Form_Load()
    Me.Move 0, 0, 11985, 9660

    Call SetOperate(Me)
    cmdDelete.Picture = LoadResPicture("CANCEL", vbResIcon)

    dtpDate(0) = Now
    dtpDate(1) = Now

    Me.Show

    chkSearch(0) = vbChecked

    Call InitGrid
    Call FillCombo

    m_nSelected = 0
End Sub

Private Sub cmdCheck_Click(Index As Integer)
    Dim i%, nValue%

    If Index = 0 Then
        nValue = flexChecked
        m_nSelected = grdData.Rows - grdData.FixedRows
    Else
        nValue = flexUnchecked
        m_nSelected = 0
    End If

    With grdData
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, i, 1) = nValue
        Next i
    End With
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then
        If chkSearch(0) = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True

            dtpDate(0).SetFocus
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False

            cmdSearch.SetFocus
        End If
    Else
        If chkSearch(Index) = vbChecked Then
            cboSearch(Index).Enabled = True

            cboSearch(Index).SetFocus
        Else
            cboSearch(Index).Enabled = False

            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus
End Sub

Private Sub cmdSearch_Click()
    Call FillGrid
End Sub

Private Sub grdData_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub

    With grdData
        If .Rows = .FixedRows Or .Row < .FixedRows Or .Row >= .Rows Then Exit Sub
    End With

    Call CheckCount
End Sub

Private Sub grdData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdData
        If .Rows = .FixedRows Or .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub
    End With

    Call CheckCount
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    grdData.ColWidth(5) = IIf(tabMain.Tab = 0, 0, 1095)

    Call FillGrid
End Sub

Private Sub cmdDelete_Click()
    Dim oLog      As PlusLib2.CLog
    Dim logData() As PlusLib2.TLOG
    Dim i&, iLog&, bReturn As Boolean

    If grdData.Rows = grdData.FixedRows Then Exit Sub

    If m_nSelected <= 0 Then Call MessageBox("선택한 내용이 없습니다. 선택 후 삭제하십시오.")

    On Error GoTo ErrHandler

    ReDim logData(m_nSelected - 1)

    With grdData
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                If .Cell(flexcpChecked, i, 1) = flexChecked Then
                    logData(iLog).LogID = CLng(.TextMatrix(i, 7))
                    logData(iLog).LogSeq = CInt(.TextMatrix(i, 8))

                    iLog = iLog + 1
                End If
            End If
        Next i
    End With

    Set oLog = New PlusLib2.CLog
    oLog.Connection = g_adoCon
    oLog.UserName = g_sUserName

    bReturn = oLog.DeleteLog(IIf(tabMain.Tab = 0, False, True), logData)

    Set oLog = Nothing

    With grdData
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then .RowHidden(i) = True
        Next i
    End With

    m_nSelected = 0

    Exit Sub

ErrHandler:
    Set oLog = Nothing

    Call ErrorBox(Err.Numbere, Err.Source, Err.Description)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    With grdData
        .Cols = 9
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .Rows = .FixedRows

        .TextArray(0) = "":
        .TextArray(1) = "선택":         .ColWidth(1) = 300
        .TextArray(2) = "로그일자":     .ColWidth(2) = 1095:        .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "컴퓨터":       .ColWidth(3) = 1095:        .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "사용자":       .ColWidth(4) = 1095:        .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "에러번호":     .ColWidth(5) = 1095:        .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "로그내용":     .ColWidth(6) = 15:          .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "LogID":        .ColWidth(7) = 0
        .TextArray(8) = "LogSeq":       .ColWidth(8) = 0

        .ColDataType(1) = flexDTBoolean
        .RowHeightMin = 405

        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FillCombo()
    Dim oLog As PlusLib2.CLog
    Dim rs   As ADODB.Recordset

    On Error GoTo ErrHandler

    Set oLog = New PlusLib2.CLog
    oLog.Connection = g_adoCon

    Set rs = oLog.GetComputerID
    With cboSearch(1)
        .Clear

        Do Until rs.EOF
            .AddItem rs!ComputerID

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If .ListCount > 0 Then .ListIndex = 0
    End With

    Set rs = oLog.GetUserID
    With cboSearch(2)
        .Clear

        Do Until rs.EOF
            .AddItem rs!UserID

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        If .ListCount > 0 Then .ListIndex = 0
    End With

    Set oLog = Nothing

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oLog = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub FillGrid()
    Dim oLog As PlusLib2.CLog
    Dim rs   As ADODB.Recordset
    Dim i&

    Screen.MousePointer = vbHourglass

    proProgress.Value = 0
    pnlProgress.Visible = True
    lblCount = LoadResString(120)

    On Error GoTo ErrHandler

    Set oLog = New PlusLib2.CLog
    oLog.Connection = g_adoCon
    Set rs = oLog.GetLog(IIf(tabMain.Tab = 0, False, True), _
        IIf(chkSearch(0) = vbChecked, 1, 0), MakeDate(DF_LONG, dtpDate(0)), MakeDate(DF_LONG, dtpDate(1) + 1), _
        IIf(chkSearch(1) = vbChecked, 1, 0), cboSearch(1), IIf(chkSearch(2) = vbChecked, 1, 0), cboSearch(2))
    Set oLog = Nothing

    With grdData
        .Redraw = flexRDNone

        .Rows = .FixedRows

        If tabMain.Tab = 0 Then
            For i = 1 To rs.RecordCount
                DoEvents

                .AddItem CStr(i) & vbTab & vbTab & rs!LogDate & vbTab & rs!ComputerID & vbTab & _
                rs!UserID & vbTab & vbTab & rs!logData & vbTab & rs!LogID & vbTab & rs!LogSeq

                lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
                proProgress.Value = CInt((i / rs.RecordCount) * 100)

                rs.MoveNext
            Next i
        Else
            For i = 1 To rs.RecordCount
                DoEvents

                .AddItem CStr(i) & vbTab & vbTab & rs!LogDate & vbTab & rs!ComputerID & vbTab & _
                rs!UserID & vbTab & rs!LogSeq & vbTab & rs!logData & vbTab & rs!LogID & vbTab & rs!LogSeq

                lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
                proProgress.Value = CInt((i / rs.RecordCount) * 100)

                rs.MoveNext
            Next i
        End If
        rs.Close
        Set rs = Nothing

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

    pnlProgress.Visible = False

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    pnlProgress.Visible = False

    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oLog = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
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
End Sub
