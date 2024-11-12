VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmCode 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "코드 찾기"
   ClientHeight    =   5790
   ClientLeft      =   5445
   ClientTop       =   4950
   ClientWidth     =   6810
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel pnlEdit 
      Height          =   885
      Left            =   1320
      TabIndex        =   9
      Top             =   2130
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1561
      _Version        =   196609
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtCode 
         Height          =   300
         Left            =   1215
         TabIndex        =   11
         Top             =   105
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   1215
         TabIndex        =   10
         Top             =   480
         Width           =   3150
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "코   드"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "명   칭"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdSelect 
      Height          =   690
      Left            =   3360
      TabIndex        =   8
      Top             =   5055
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      선택(&Q)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   510
      Left            =   4320
      TabIndex        =   1
      Top             =   4215
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
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   5100
      TabIndex        =   0
      Top             =   5055
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   900
      Left            =   30
      TabIndex        =   2
      Top             =   45
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1588
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   2685
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   18
         ToolTipText     =   "자료 저장"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   4275
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   17
         ToolTipText     =   "자료 추가"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   5865
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   16
         ToolTipText     =   "자료 삭제"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   5070
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   15
         ToolTipText     =   "자료 수정"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   3480
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   14
         ToolTipText     =   "자료 취소"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   75
         TabIndex        =   3
         Top             =   465
         Width           =   2025
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   90
         TabIndex        =   4
         Top             =   120
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "코드명 검색"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdAll 
         Height          =   330
         Left            =   2160
         TabIndex        =   5
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
      Height          =   3990
      Left            =   15
      TabIndex        =   7
      Top             =   990
      Width           =   6750
      _cx             =   11906
      _cy             =   7038
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
      Left            =   270
      TabIndex        =   6
      Top             =   5310
      Width           =   945
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const LIMIT_WIDTH = 3860 '3140
Private Const LIMIT_ROW = 16

Private m_sFlag        As String * 1
Private m_bSortForward As Boolean


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

Private Sub ClearData()
    txtCode = ""
    txtName = ""
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim bResult As Boolean
        
    On Error GoTo ErrHandler
    '---------------------------------------------------------------------------
    Select Case Index   '[1] 추가
        Case ID_ADDNEW
            m_sFlag = ID_ADDNEW
            Call ChangeMode(Me, False)
            tabForm.Enabled = False
            Call ClearData
            pnlMsg.Caption = LoadResString(302)
            
            txtCode.SetFocus
    '---------------------------------------------------------------------------
        Case ID_UPDATE '[2] 수정
            m_sFlag = ID_UPDATE
            Call ChangeMode(Me, False)
            
            pnlMsg.Caption = LoadResString(303)
            tabForm.Enabled = False
            txtCode.Locked = True
            txtName.SetFocus
    '---------------------------------------------------------------------------
        Case ID_DELETE '[3] 삭제
            If grdData.Rows = grdData.FixedRows Then Exit Sub
    
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                m_sFlag = ID_DELETE
                
                If SaveData() Then Call FillGrid
            End If
    '---------------------------------------------------------------------------
        Case ID_SAVE  '[4] 저장
            If CheckData() = False Then Exit Sub
            If SaveData() Then
                Call FillGrid
                Call ChangeMode(Me, True)
                
                m_sFlag = ""
                tabForm.Enabled = True
                txtCode.Locked = False
            End If
    '---------------------------------------------------------------------------
        Case ID_CANCEL '[5] 취소
            m_sFlag = ""
            If grdData.Rows > 1 Then
                Call ShowData
            Else
                Call ClearData
            End If
            Call ChangeMode(Me, True)
            tabForm.Enabled = True
            txtCode.Locked = False
    End Select
    
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "Code.cmdOperate_Click", Err.Description)
End Sub

Private Sub ShowData()
    With grdData
        txtCode = .TextMatrix(.Row, 1)
        txtName = .TextMatrix(.Row, 2)
    End With
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 9555, 7500
    
    Call SetOperate(Me)
    
    Call InitGrid
    Call FillGrid
    
    cmdAll.Picture = LoadResPicture("ALL", vbResIcon)
    lblCount.Caption = LoadResString(250)
End Sub

Private Sub InitGrid()
    Call SetVSFlexGrid(grdData)
    With grdData
        .Redraw = False
        .Rows = 1
        .Cols = 3
        
        .TextArray(0) = ""
        .TextArray(1) = "코드": .ColWidth(1) = 450: .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "부서명": .ColWidth(2) = LIMIT_WIDTH: .ColAlignment(2) = flexAlignLeftCenter
        
        .Redraw = True
    End With
End Sub

Private Sub grdData_DblClick()
    With grdData
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub

        Call cmdOperate_Click(ID_UPDATE)
    End With
End Sub

Private Sub grdData_RowColChange()
    Call ShowData
End Sub

Private Sub tabForm_Click(PreviousTab As Integer)
    Dim sMenuID As String
    
    pnlCaption(0).Caption = Replace(tabForm.TabCaption(tabForm.Tab), "관리", "")
    pnlCaption(3).Caption = Replace(tabForm.TabCaption(tabForm.Tab), "관리", "") & "검색"
    grdData.TextMatrix(0, 2) = Replace(tabForm.TabCaption(tabForm.Tab), "관리", "")

    Call WizMDI.RunForm(1210 + (10 * tabForm.Tab))

    Call FillGrid
End Sub

Private Sub FillGrid()
    Dim oCode As WizLib.CCode
    Dim rs As ADODB.Recordset
    Dim lNowRow&
    
    On Error GoTo ErrHandler
        
    Set oCode = New WizLib.CCode
    oCode.Connection = g_adoCon
        
    Select Case tabForm.Tab
        Case 0 ' [0] 원단폭
            oCode.CodeType = CD_WIDTH
        Case 1 ' [1] 가공구분 관리
            oCode.CodeType = CD_WORK
        Case 2 ' [2] 레벨구분 관리
            oCode.CodeType = CD_LABEL
        Case 3 ' [3] 밴드구분 관리
            oCode.CodeType = CD_BAND
        Case 4 ' [5] 주문형태 관리
            oCode.CodeType = CD_FORM
        Case 5 ' [6] 주문구분 관리
            oCode.CodeType = CD_CLASS
    End Select
    Set rs = oCode.GetCode()
    Set oCode = Nothing
    
    If rs.RecordCount = 0 Then
        grdData.Rows = grdData.FixedRows
        grdData.HighLight = flexHighlightNever
        lblCount.Caption = LoadResString(250)
        
        rs.Close
        Set rs = Nothing
        
        Call ClearData
        Call ChangeScroll
        Exit Sub
    End If
    
    With grdData
        .Redraw = False
        
        lNowRow = IIf(.Row > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs.Fields(0) & vbTab & rs.Fields(1)
            rs.MoveNext
        Loop
    
        Call ChangeScroll
        
        lblCount.Caption = LoadResString(250) & grdData.Rows - 1 & " 건"
        
        rs.Close
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
           .HighLight = flexHighlightAlways
           .Row = IIf(.Rows > lNowRow, lNowRow, .Rows - 1)
           .TopRow = lNowRow
           
           .Col = .FixedCols
           .ColSel = .Cols - 1
           
            Call ShowData
        End If
        .Redraw = True
    End With
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCode = Nothing
    
    Call ErrorBox(Err.Number, "Code.FillGrid", Err.Description)
End Sub

Private Function CheckData() As Boolean
    Dim i%
    
    CheckData = True
    
    If m_sFlag = ID_ADDNEW Or m_sFlag = ID_UPDATE Then
        If Len(txtName) = 0 Then
            MsgBox LoadResString(115), vbInformation
            txtName.SetFocus
            CheckData = False
            Exit Function
        End If
    End If
End Function

Private Sub ChangeScroll()
    With grdData
        If .Rows > LIMIT_ROW Then
            .ColWidth(2) = LIMIT_WIDTH - 240
        Else
            .ColWidth(2) = LIMIT_WIDTH
        End If
    End With
End Sub

Private Function SaveData() As Boolean
    Dim NewCode As WizLib.tCode
    Dim oCode As WizLib.CCode
    
    On Error GoTo ErrHandler
    
    Set oCode = New WizLib.CCode
    oCode.Connection = g_adoCon
    oCode.UserName = g_sUserName
    
    NewCode.sCodeID = Format(txtCode, "0#")
    NewCode.scode = txtName
    
    Select Case tabForm.Tab
        Case 0 ' [0] 원단폭
            oCode.CodeType = CD_WIDTH
        Case 1 ' [1] 가공구분 관리
            oCode.CodeType = CD_WORK
        Case 2 ' [2] 레벨구분 관리
            oCode.CodeType = CD_LABEL
        Case 3 ' [3] 밴드구분 관리
            oCode.CodeType = CD_BAND
        Case 4 ' [5] 주문형태 관리
            oCode.CodeType = CD_FORM
        Case 5 ' [6] 주문구분 관리
            oCode.CodeType = CD_CLASS
      
    End Select
    
    If m_sFlag = ID_ADDNEW Then
        SaveData = oCode.AddNewCode(NewCode)
    ElseIf m_sFlag = ID_UPDATE Then
        SaveData = oCode.UpdateCode(NewCode)
    ElseIf m_sFlag = ID_DELETE Then
        SaveData = oCode.DeleteCode(NewCode.sCodeID)
    End If
    
    Set oCode = Nothing
    Exit Function
ErrHandler:
    Set oCode = Nothing
    Call ErrorBox(Err.Number, "Code.SaveData", Err.Description)
End Function




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

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

