VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOutCode 
   Caption         =   "출고관련 코드관리"
   ClientHeight    =   6720
   ClientLeft      =   6210
   ClientTop       =   1800
   ClientWidth     =   8655
   Icon            =   "frmOutCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   8655
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   6960
      TabIndex        =   11
      Top             =   6000
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   4815
      Left            =   60
      TabIndex        =   19
      Top             =   1035
      Width           =   3990
      _cx             =   7038
      _cy             =   8493
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
   Begin Threed.SSPanel pnlMsg 
      Height          =   510
      Left            =   4305
      TabIndex        =   13
      Top             =   5100
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
   Begin Threed.SSPanel pnlEdit 
      Height          =   840
      Left            =   4125
      TabIndex        =   12
      Top             =   1050
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   1482
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Left            =   1185
         TabIndex        =   6
         Top             =   450
         Width           =   3120
         _ExtentX        =   5503
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
         Left            =   1185
         TabIndex        =   5
         Top             =   90
         Width           =   1110
         _ExtentX        =   1958
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
         BackColor       =   12648384
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   90
         TabIndex        =   0
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "코     드"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   450
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "출고 구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin TabDlg.SSTab tabForm 
      Height          =   5340
      Left            =   15
      TabIndex        =   9
      Top             =   975
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   9419
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   679
      TabCaption(0)   =   "출고구분 관리 "
      TabPicture(0)   =   "frmOutCode.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "반품구분 관리"
      TabPicture(1)   =   "frmOutCode.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin Threed.SSPanel pnlBoard 
      Height          =   915
      Left            =   4080
      TabIndex        =   10
      Top             =   30
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   1614
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   1350
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   8
         ToolTipText     =   "자료 취소"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   2940
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   3
         ToolTipText     =   "자료 수정"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   3735
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   4
         ToolTipText     =   "자료 삭제"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   2145
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   2
         ToolTipText     =   "자료 추가"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   555
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   7
         ToolTipText     =   "자료 저장"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   900
      Left            =   30
      TabIndex        =   14
      Top             =   45
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   1588
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   75
         TabIndex        =   15
         Top             =   465
         Width           =   2025
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   90
         TabIndex        =   16
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
      Left            =   135
      TabIndex        =   18
      Top             =   6450
      Width           =   945
   End
End
Attribute VB_Name = "frmOutCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LIMIT_WIDTH = 3140
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
            
            grdData.SetFocus
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
            grdData.SetFocus
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

Private Sub cmdPrint_Click()

End Sub

Private Sub Form_Activate()
    txtSearch.SetFocus
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 8775, 7125
    
    Call SetOperate(Me)
    
    m_bFirst = True
    
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

Private Sub grdData_KeyDown(KeyCode As Integer, Shift As Integer)
    With grdData
    If .Rows > .FixedRows Then
        If KeyCode = vbKeyReturn Then
            Call cmdOperate_Click(ID_UPDATE)
        End If
    End If
    End With
End Sub

Private Sub grdData_RowColChange()
    Call ShowData
End Sub

Private Sub tabForm_Click(PreviousTab As Integer)
    pnlCaption(0).Caption = Replace(tabForm.TabCaption(tabForm.Tab), "관리", "")
    pnlCaption(3).Caption = Replace(tabForm.TabCaption(tabForm.Tab), "관리", "") & "검색"
    grdData.TextMatrix(0, 2) = Replace(tabForm.TabCaption(tabForm.Tab), "관리", "")

    Call PlusMDI.RunForm(1410 + (10 * tabForm.Tab))

    Call FillGrid
    txtSearch.SetFocus
End Sub

Private Sub FillGrid()
    Dim oCode As PlusLib2.CCode
    Dim rs As ADODB.Recordset
    Dim lNowRow&
    
    On Error GoTo ErrHandler
        
    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon
        
    Select Case tabForm.Tab
        Case 0 ' [0] 출고구분 관리
            oCode.CodeType = CD_OUTCLSS
        Case 1 ' [1] 반품구분 관리
            oCode.CodeType = CD_BACKCLSS
    End Select
    Set rs = oCode.Getcode()
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
    Dim NewCode As PlusLib2.tCode
    Dim oCode As PlusLib2.CCode
    

    On Error GoTo ErrHandler
    
    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon
    oCode.UserName = g_sUserName
    
    NewCode.sCodeID = Format(txtCode, "0#")
    NewCode.scode = txtName
    
    Select Case tabForm.Tab
        Case 0  '[0] 출고구분 관리
            oCode.CodeType = CD_OUTCLSS
        Case 1  '[1] 반품구분 관리
            oCode.CodeType = CD_BACKCLSS
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
    Call ChangeMode(Me, True)
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
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        grdData.SetFocus
    End If
End Sub

