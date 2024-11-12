VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmColor 
   Caption         =   "색상명관리"
   ClientHeight    =   7050
   ClientLeft      =   3525
   ClientTop       =   5220
   ClientWidth     =   7605
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   7605
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   5265
      Left            =   15
      TabIndex        =   21
      Top             =   975
      Width           =   3360
      _cx             =   5927
      _cy             =   9287
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
   Begin Crystal.CrystalReport cryReport 
      Left            =   2985
      Top             =   6405
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlBoard 
      Height          =   6195
      Left            =   3420
      TabIndex        =   6
      Top             =   45
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   10927
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel pnlEdit 
         Height          =   840
         Left            =   45
         TabIndex        =   7
         Top             =   915
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   1482
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MRPPlus2.WizText txtCode 
            Height          =   300
            Left            =   1140
            TabIndex        =   13
            Top             =   75
            Width           =   960
            _ExtentX        =   1693
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
            MaxLength       =   4
            BackColor       =   12648384
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Left            =   1140
            TabIndex        =   14
            Top             =   420
            Width           =   2835
            _ExtentX        =   5001
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
            Left            =   75
            TabIndex        =   8
            Top             =   75
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   " 색상코드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   75
            TabIndex        =   9
            Top             =   435
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "색 상  명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   900
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   16
         ToolTipText     =   "자료 취소"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   2490
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
         Left            =   3285
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
         Left            =   1695
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   10
         ToolTipText     =   "자료 추가"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   120
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   15
         ToolTipText     =   "자료 저장"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   510
         Left            =   390
         TabIndex        =   20
         Top             =   5415
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   900
         _Version        =   196609
         BackColor       =   65535
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   1588
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optSize 
         Caption         =   "상세"
         Height          =   330
         Index           =   1
         Left            =   2655
         Style           =   1  '그래픽
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton optSize 
         Caption         =   "요약"
         Height          =   330
         Index           =   0
         Left            =   2655
         Style           =   1  '그래픽
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   645
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   75
         TabIndex        =   2
         Top             =   465
         Width           =   2025
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   25
         Left            =   90
         TabIndex        =   1
         Top             =   120
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "색상명 검색"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdAll 
         Height          =   330
         Left            =   2130
         TabIndex        =   3
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
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   4170
      TabIndex        =   17
      Top             =   5490
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   5910
      TabIndex        =   18
      Top             =   6330
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
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
      Height          =   180
      Left            =   105
      TabIndex        =   19
      Top             =   6480
      Width           =   945
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const REPORTFILE = "\Report\Article.rpt"

Private Const LIMIT_ROW = 14
Private Const LIMIT_WIDTH = 2400

Dim m_sFlag        As String * 1
Dim m_bSortForward As Boolean
Dim m_bSkip As Boolean

Private Sub cmdAll_Click()
    Dim iLoop As Integer

    With grdData
        .Redraw = flexRDNone

        For iLoop = .FixedRows To .Rows - .FixedRows
            .RowHidden(iLoop) = False
        Next iLoop

        .Redraw = flexRDDirect
    End With

    txtSearch.Text = ""
    cmdAll.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOperate_Click(Index As Integer)

    On Error GoTo ErrHandler
    
    If optSize(0).Value Then optSize(1).Value = True

    Select Case Index
        Case ID_ADDNEW
            m_sFlag = ID_ADDNEW
            Call ClearData
            Call ChangeMode(Me, False)
            txtCode.Locked = False
            txtName.SetFocus
            pnlMsg.Caption = LoadResString(302)
            
        Case ID_UPDATE
            m_sFlag = ID_UPDATE
            Call ChangeMode(Me, False)
            txtCode.Locked = True
            txtName.SetFocus
            pnlMsg.Caption = LoadResString(303)
            
        Case ID_DELETE
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                m_sFlag = ID_DELETE
                If SaveData Then
                    Call FillGrid
                    m_sFlag = ""
                End If
            End If
            
        Case ID_SAVE
            If Not CheckData() Then Exit Sub
    
            If SaveData() Then
                Call ChangeMode(Me, True)
                Call FillGrid
                m_sFlag = ""
            End If
            grdData.SetFocus
        Case ID_CANCEL
            m_sFlag = ""
            Call ChangeMode(Me, True)
            With grdData
                If .Rows > .FixedRows Then
                    Call ShowData
                Else
                    Call ClearData
                End If
            End With
            grdData.SetFocus
            
        End Select

    Exit Sub
ErrHandler:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Err.Clear
End Sub


Private Function CheckData() As Boolean
    Dim i%
    CheckData = True
    If m_sFlag = ID_ADDNEW Then
        With grdData
            For i = 1 To .Rows - 1
                If Trim(txtCode) = .TextMatrix(i, 1) Then
                    MsgBox LoadResString(114), vbInformation
                    txtCode.SetFocus
                    CheckData = False
                    Exit Function
                End If
            Next i
        End With
    End If
    
    If Len(txtName) = 0 Then
        MsgBox "색상명이 없습니다. 색상명을 넣어 주십시오", vbInformation
        txtName.SetFocus
        CheckData = False
        Exit Function
    End If

End Function

Private Function SaveData() As Boolean
    Dim oColor As PlusLib2.CColor
    Dim NewColor As PlusLib2.TColor
    
    On Error GoTo ErrHandler
    
    With NewColor
        .sColorID = IIf(Len(txtCode) > 0, Format(txtCode, "00000"), "")
        .sColor = txtName
    End With
    
    Set oColor = New PlusLib2.CColor
    oColor.Connection = g_adoCon
    oColor.UserName = g_sUserName
    
    Select Case m_sFlag
        Case ID_ADDNEW
            SaveData = oColor.AddNewColor(NewColor)
        Case ID_UPDATE
            NewArticle.sArticleID = grdData.TextMatrix(grdData.Row, 1)
            SaveData = oColor.UpdateColor(NewColor)
        Case ID_DELETE
            SaveData = oColor.DeleteColor(grdData.TextMatrix(grdData.Row, 1))
    End Select
    
    Set oColor = Nothing
    Exit Function
ErrHandler:
    Set oColor = Nothing

    Call ErrorBox(Err.Number, "frmColor.SaveData", Err.Description)
End Function


Private Sub cmdPrint_Click()
    Dim oArticle As PlusLib2.CArticle
    Dim rs As ADODB.Recordset
    Dim lNowRow&
    Dim sParam() As String
    
    On Error GoTo ErrHandler
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass

    Set oArticle = New PlusLib2.CArticle
    oArticle.Connection = g_adoCon
    
    Set rs = oArticle.GetArticle(IIf(Len(txtSearch) = 0, "", "%" & txtSearch))
    Set oArticle = Nothing
    
    ReDim sParam(2)
    sParam(0) = "품목 리스트"
    sParam(1) = "검색조건 : " & IIf(Len(txtSearch.Text) > 0, txtSearch, "(전체)")
    sParam(2) = CompanyName
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.PrintPreview)
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "Article.cmdPrint_Click", Err.Description)
End Sub

Private Sub Form_Activate()
    txtSearch.SetFocus
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 7725, 6165
    
    Call SetOperate(Me)
    
    Call InitGrid
    Call FillGrid
    
    txtCode.MaxLength = 5
    cmdAll.Picture = LoadResPicture("ALL", vbResIcon)
End Sub

Private Sub InitGrid()
    Call SetVSFlexGrid(grdData)
    With grdData
        .Redraw = False
        .Rows = 1
        .Cols = 3
        
        .TextArray(0) = ""
        .TextArray(1) = "코드":     .ColWidth(1) = 570:             .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "색상명":     .ColWidth(2) = LIMIT_WIDTH:     .ColAlignment(2) = flexAlignLeftCenter
        
        .Redraw = True
    End With
End Sub

Private Sub FillGrid()
    Dim oColor As PlusLib2.CColor
    Dim rs As ADODB.Recordset
    Dim lNowRow&
    
    On Error GoTo ErrHandler
    
    m_bSkip = True
    
    Set oColor = New PlusLib2.CColor
    oColor.Connection = g_adoCon
    
    Set rs = oColor.AddNewArticle
    Set oArticle = Nothing
    
    If rs.RecordCount = 0 Then
        grdData.Rows = grdData.FixedRows
        grdData.HighLight = flexHighlightNever
        lblCount.Caption = LoadResString(250)
        
        Call ClearData
        Call ChangeScroll
        Exit Sub
    End If
    
    With grdData
        .Redraw = False
        If .Rows > .FixedRows Then
            If m_sFlag = ID_ADDNEW Then
                lNowRow = .Rows
            Else
                lNowRow = .Row
            End If
            .Rows = 1
        Else
            lNowRow = 1
        End If
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & CStr(rs!ColorID) & vbTab & rs!Color
            
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        Call ChangeScroll
        lblCount.Caption = LoadResString(250) & .Rows - 1 & " 건"
    
        If .Rows > .FixedRows Then
            If .Rows > lNowRow Then
                .Row = lNowRow
            Else
                .Row = .Rows - 1
            End If
            .TopRow = .Row
            .HighLight = flexHighlightAlways
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
            Call ShowData
        End If
        .Redraw = True
    End With
    
    m_bSkip = False
    Exit Sub

ErrHandler:
    Set oArticle = Nothing
    Call ErrorBox(Err.Number, "Article.FillGrid", Err.Description)
End Sub

Private Sub ChangeScroll()
    Dim lRows As Long
    
    lRows = GetVisibleVSGridRowCount(grdData)
    
    With grdData
        If lRows > LIMIT_ROW Then
            .ColWidth(2) = LIMIT_WIDTH - 240
        Else
            .ColWidth(2) = LIMIT_WIDTH
        End If
    End With

    If lRows = 0 Then
        Call ClearData
        cmdOperate(ID_UPDATE).Enabled = False
        cmdOperate(ID_DELETE).Enabled = False
        cmdPrint.Enabled = False
    Else
        Call ShowData
        cmdOperate(ID_UPDATE).Enabled = True
        cmdOperate(ID_DELETE).Enabled = True
        cmdPrint.Enabled = True
    End If
End Sub

Private Sub ClearData()
    txtCode = ""
    txtName = ""
End Sub

Private Sub ShowData()
    On Error Resume Next
    
    With grdData
        txtCode = .TextMatrix(.Row, 1)  '[코드]
        txtName = .TextMatrix(.Row, 2)  '[1] 색상명
    End With
End Sub

Private Sub grdData_AfterSort(ByVal Col As Long, Order As Integer)
    Call ShowData
End Sub

Private Sub grdData_DblClick()
    With grdData
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub
    End With
    
    Call cmdOperate_Click(ID_UPDATE)
End Sub

Private Sub grdData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdOperate_Click(ID_UPDATE)
    End If
End Sub

Private Sub grdData_RowColChange()
    If m_bSkip Then Exit Sub
    Call ShowData
End Sub

Private Sub optSize_Click(Index As Integer)
    Dim lRows As Long
    
    lRows = GetVisibleVSGridRowCount(grdData)
    
    If optSize(0).Value Then '[0] 요약
        With grdData
            .Width = 7625
            If lRows > LIMIT_ROW Then
                .ColWidth(2) = LIMIT_WIDTH + 1560
            Else
                .ColWidth(2) = LIMIT_WIDTH + 1800
            End If
            .ColWidth(3) = 2400
        End With
    Else '[1] 상세
        With grdData
            .Width = 3420
            If lRows > LIMIT_ROW Then
                .ColWidth(2) = LIMIT_WIDTH - 240
            Else
                .ColWidth(2) = LIMIT_WIDTH
            End If
            .ColWidth(3) = 0
        End With
    End If
End Sub

Private Sub txtSearch_Change()
    Dim iLoop  As Integer
    Dim iCount As Integer
    Dim iNowRow As Integer

    On Error GoTo ErrHandler
    
    If Len(Trim(txtSearch)) > 0 Then
        With grdData
            .Redraw = False

            For iLoop = .FixedRows To .Rows - .FixedRows
                If InStr(UCase(.TextArray(iLoop * .Cols + 2)), UCase(txtSearch)) = 0 Then
                    .RowHidden(iLoop) = True
                    iCount = iCount + 1
                Else
                    .RowHidden(iLoop) = False
                    iNowRow = iLoop
                End If
            Next iLoop

            If iNowRow > .FixedRows Then
                .Row = iNowRow
                
                .Col = .FixedCols
                .ColSel = .Cols - 1
            End If

            .Redraw = True
            .TopRow = .Row
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
    Call ErrorBox(Err.Number, "Article.txtSearch_Change", Err.Description)

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    'Call MoveFocus(KeyCode)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        grdData.SetFocus
    End If
End Sub
