VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmOrderCode 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "가공구분 찾기"
   ClientHeight    =   6615
   ClientLeft      =   4530
   ClientTop       =   1500
   ClientWidth     =   9435
   Icon            =   "frmOrderCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel pnlMsg 
      Height          =   510
      Left            =   5085
      TabIndex        =   15
      Top             =   5055
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
      Height          =   4815
      Left            =   4815
      TabIndex        =   14
      Top             =   990
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   8493
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1155
         TabIndex        =   9
         Top             =   465
         Width           =   3330
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Left            =   1155
         TabIndex        =   8
         Top             =   90
         Width           =   1020
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   75
         TabIndex        =   3
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
         Left            =   75
         TabIndex        =   4
         Top             =   465
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "가공구분"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   0
         X2              =   4590
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   0
         X2              =   4590
         Y1              =   870
         Y2              =   870
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   7695
      TabIndex        =   13
      Top             =   5880
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlBoard 
      Height          =   900
      Left            =   4815
      TabIndex        =   12
      Top             =   45
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   1588
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
         TabIndex        =   11
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   5
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
         TabIndex        =   10
         ToolTipText     =   "자료 저장"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   900
      Left            =   30
      TabIndex        =   16
      Top             =   45
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   1588
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   75
         TabIndex        =   1
         Top             =   465
         Width           =   2025
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   90
         TabIndex        =   17
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
         TabIndex        =   2
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
      Height          =   4815
      Left            =   30
      TabIndex        =   0
      Top             =   990
      Width           =   4710
      _cx             =   8308
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
   Begin Threed.SSCommand cmdSelect 
      Height          =   690
      Left            =   5910
      TabIndex        =   19
      Top             =   5880
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      선택(&Q)"
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
      TabIndex        =   18
      Top             =   6075
      Width           =   945
   End
End
Attribute VB_Name = "frmOrderCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const LIMIT_WIDTH = 3860 '3140
Private Const LIMIT_ROW = 16

Private m_bSortForward As Boolean

Dim m_nFlag         As Integer
Dim m_bSelected     As Boolean
Dim wData()

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
    m_bSelected = False
    
    Unload Me
End Sub

Private Sub ClearData()
    txtCode = ""
    txtName = ""
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    '---------------------------------------------------------------------------------------
    Case ID_ADDNEW
        m_nFlag = ID_ADDNEW
        
        Call ClearData
        Call ChangeMode(False)
        txtCode.Locked = False
        txtName.SetFocus
        pnlMsg.Caption = LoadResString(302) '자료 입력(추가) 중 ...
    '---------------------------------------------------------------------------------------
    Case ID_UPDATE
        If grdData.Rows > grdData.FixedRows Then
            m_nFlag = ID_UPDATE
            
            Call ChangeMode(False)
            txtCode.Locked = True
            txtName.SetFocus
            pnlMsg.Caption = LoadResString(303) '자료 수정 중 ...
        End If
    '---------------------------------------------------------------------------------------
    Case ID_DELETE
        If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then '선택하신 항목을 삭제하시겠습니까 ?
            m_nFlag = ID_DELETE
            If SaveData() Then
                Call SetGrid(FL_BY_NAME)
                If Len(txtSearch) > 0 Then Call txtSearch_Change
                m_nFlag = 9
            End If
        End If
    '---------------------------------------------------------------------------------------
    Case ID_SAVE
        If Not CheckData() Then Exit Sub

        If SaveData() Then
            Call ChangeMode(True)
            Call SetGrid(FL_BY_NAME)
            If Len(txtSearch) > 0 Then Call txtSearch_Change
            
            m_nFlag = ID_CANCEL
        End If
        grdData.SetFocus
    '---------------------------------------------------------------------------------------
    Case ID_CANCEL
        m_nFlag = ID_CANCEL
        Call ChangeMode(True)
        With grdData
            If .Rows > .FixedRows Then
                Call ShowData
                .SetFocus
            Else
                Call ClearData
                txtSearch.SetFocus
            End If
        End With
    End Select
End Sub

Private Sub ShowData()
    With grdData
        txtCode = .TextMatrix(.Row, 1)
        txtName = .TextMatrix(.Row, 2)
    End With
End Sub

Private Sub cmdSelect_Click()
    If grdData.Rows = grdData.FixedRows Then Exit Sub

    Call SelectData

End Sub

Private Sub Form_Load()
   
    m_nFlag = ID_CANCEL
    
    Call SetOperate
    
    Call InitGrid
   
    txtCode.MaxLength = 2
    cmdAll.Picture = LoadResPicture("ALL", vbResIcon)
'    lblCount.Caption = LoadResString(250)
End Sub

Private Sub InitGrid()
    With grdData
        .Redraw = flexRDNone
        
        .Rows = 1
        .RowHeight(0) = 450
        .ColWidth(0) = 360

        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .FillStyle = flexFillRepeat
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        .AllowSelection = False
        .AllowBigSelection = False
        .ExtendLastCol = True
        
        .Editable = flexEDNone
        .MousePointer = flexCustom

        .RowHeightMin = 275
        .WordWrap = True

        .ColAlignment(0) = flexAlignCenterCenter
        For iCount = .FixedCols To .Cols - 1
            .FixedAlignment(iCount) = flexAlignCenterCenter
        Next iCount
        
        .Rows = 1
        .Cols = 3
        
        .TextArray(0) = ""
        .TextArray(1) = "코드": .ColWidth(1) = 450: .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "가공구분": .ColWidth(2) = LIMIT_WIDTH: .ColAlignment(2) = flexAlignLeftCenter
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmOrderCode = Nothing
End Sub

Private Sub grdData_AfterSort(ByVal Col As Long, Order As Integer)
    Call ShowData
End Sub

Private Sub grdData_DblClick()
    With grdData
        If .Row < 1 Or .Row >= .Rows Then Exit Sub

        Call SelectData
    End With
End Sub

Private Sub grdData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call grdData_DblClick
    End If
End Sub

Private Sub grdData_RowColChange()
    If m_bSkip Then Exit Sub
    
    Call ShowData
End Sub

Private Function CheckData() As Boolean
    Dim i%
    CheckData = True
    If m_nFlag = ID_ADDNEW Then
        With grdData
            For i = 1 To .Rows - 1
                If Trim(txtCode) = .TextMatrix(i, 1) Then
                    MsgBox LoadResString(114), vbInformation '이미 등록된 코드가 있습니다. 다른 코드를 넣어주십시오.
                    txtCode.SetFocus
                    CheckData = False
                    Exit Function
                End If
            Next i
        End With
    End If
    
    If Len(txtName) = 0 Then
        MsgBox "가공구분명이 없습니다. 가공구분명을 넣어 주십시오", vbInformation
        txtName.SetFocus
        CheckData = False
        Exit Function
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
    Dim NewCode As PlusLib2.TCode
    Dim oCode As PlusLib2.CCode
    
    On Error GoTo ErrHandler
    
    Set oCode = New PlusLib2.CCode
    oCode.Connection = adoCon
'    oCode.UserName = g_sUserName
    
    NewCode.sCodeID = Format(txtCode, "0#")
    NewCode.scode = txtName
    oCode.CodeType = CD_WORK
    
    If m_nFlag = ID_ADDNEW Then
        SaveData = oCode.AddNewCode(NewCode)
    ElseIf m_nFlag = ID_UPDATE Then
        SaveData = oCode.UpdateCode(NewCode)
    ElseIf m_nFlag = ID_DELETE Then
        SaveData = oCode.DeleteCode(NewCode.sCodeID)
    End If
    
    Set oCode = Nothing
    Exit Function
ErrHandler:
    Set oCode = Nothing
    Call ErrorBox(Err.Number, "Code.SaveData", Err.Description)
End Function


Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtName_Change()
    With txtName
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
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
    ElseIf KeyCode = vbKeyReturn Then
        Call SetGrid(FL_BY_NAME, txtSearch)
    End If
End Sub

Public Function SetMsg(SelData(), Optional sNewData) As Boolean
    Dim i%
       
       
    If IsMissing(sNewData) Then
        Me.Show vbModal
    Else
        If sNewData = "" Then
            Me.Show vbModal
        Else
            Call SetGrid(FL_BY_CODE, sNewData)
            If grdData.Rows = grdData.FixedRows Then
                txtSearch = sNewData
                Call SetGrid(FL_BY_NAME, sNewData)
            End If
            
            '------------------------------------------------
            With grdData
                If .Rows > .FixedRows Then
                    If .Rows = .FixedRows + 1 Then
                        Call SelectData
                    Else
                        Me.Show vbModal
                    End If
                Else
                    If MsgBox(LoadResString(112), vbQuestion + vbYesNo) = vbYes Then
                        Call cmdOperate_Click(ID_ADDNEW)
                        txtName.Text = sNewData
                        
                        Me.Show vbModal
                    Else
                        Me.Show vbModal
                    End If
                End If
            End With
        End If
    End If
    
    '=====================================================================
    If m_bSelected Then
        With grdData
            ReDim SelData(UBound(wData) - 1)
            For i = LBound(wData) To UBound(wData) - 1
                SelData(i) = wData(i)
            Next i
        End With
    End If
    
    SetMsg = m_bSelected
End Function


Private Sub SetGrid(ByVal Index As EFindClss, Optional sNewData)
    Dim oCode As PlusLib2.CCode
    Dim rs As ADODB.Recordset
    
    Dim nNowRow&, sID$
    
    On Error GoTo ErrHandler
    
    m_bSkip = True
       
    Set oCode = New PlusLib2.CCode
    oCode.Connection = adoCon
    oCode.CodeType = CD_WORK
        
    If Index = FL_BY_CODE Then
        If LenB(StrConv(sNewData, vbFromUnicode)) < 2 Then
            Set rs = oCode.GetcodeID(CStr(sNewData))
        Else
            Set oCode = Nothing
            Exit Sub
        End If
    ElseIf Index = FL_BY_NAME Then
        Set rs = oCode.GetcodeOne(sNewData)
    End If
    Set oArticle = Nothing
    
    With grdData
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            If m_nFlag = ID_ADDNEW Then
                nNowRow = .Rows
            Else
                nNowRow = .Row
            End If
            .Rows = .FixedRows
        Else
            nNowRow = 1
        End If
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs!WorkID & vbTab & _
                    rs!Work
        
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        lblCount.Caption = "검색건수 : " & CStr(.Rows - 1) & " 건"
        
        If .Rows > .FixedRows Then
            If .Rows > nNowRow Then
                .Row = nNowRow
            Else
                .Row = .Rows - 1
            End If
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
            .HighLight = flexHighlightAlways
            
            Call ShowData
        Else
            .HighLight = flexHighlightNever
            
            Call ClearData
        End If
        
        .Redraw = flexRDDirect
    End With
    
    m_bSkip = False
    Exit Sub
ErrHandler:
    Set oArticle = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmOrderCode.SetGrid", Err.Description)
End Sub

Private Sub SelectData()
    Dim i%
    
    On Error Resume Next
    
    If grdData.Rows > 1 Then
        m_bSelected = True
        
        ReDim wData(grdData.Cols - 1)
        With grdData
            For i = 1 To .Cols - 1
                wData(i - 1) = .TextMatrix(.Row, i)
            Next i
        End With
        
        Me.Hide
    End If
End Sub

Private Sub SetOperate()
    Dim oControl As Object

    For Each oControl In Me.Controls
        If (TypeOf oControl Is SSCommand) Or (TypeOf oControl Is CommandButton) _
            Or (TypeOf oControl Is SSOption) Or (TypeOf oControl Is OptionButton) Then
            oControl.MousePointer = ssCustom
            oControl.MouseIcon = LoadResPicture("POINTER", vbResCursor)
        End If
    Next oControl

    pnlEdit.Enabled = False

    cmdOperate(ID_SAVE).Visible = False
    cmdOperate(ID_CANCEL).Visible = False
    cmdExit.Cancel = True

    cmdOperate(ID_ADDNEW).Picture = LoadResPicture("ADDNEW", vbResIcon)
    cmdOperate(ID_UPDATE).Picture = LoadResPicture("UPDATE", vbResIcon)
    cmdOperate(ID_DELETE).Picture = LoadResPicture("DELETE", vbResIcon)
    cmdOperate(ID_SAVE).Picture = LoadResPicture("SAVE", vbResIcon)
    cmdOperate(ID_CANCEL).Picture = LoadResPicture("CANCEL", vbResIcon)

    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
    cmdSelect.Picture = LoadResPicture("SELECT", vbResIcon)
End Sub

Private Sub ChangeMode(bFlag As Boolean)
    On Error Resume Next

    pnlEdit.Enabled = Not bFlag
    pnlSearch.Enabled = bFlag
    pnlMsg.Visible = Not bFlag

    grdData.Enabled = bFlag

    cmdOperate(ID_ADDNEW).Enabled = bFlag
    cmdOperate(ID_UPDATE).Enabled = bFlag
    cmdOperate(ID_DELETE).Enabled = bFlag

    cmdOperate(ID_SAVE).Visible = Not bFlag
    cmdOperate(ID_CANCEL).Visible = Not bFlag

    If bFlag Then
        cmdExit.Cancel = True
    Else
        cmdOperate(ID_CANCEL).Cancel = True
    End If
End Sub

