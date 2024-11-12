VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDyeAux 
   ClientHeight    =   6705
   ClientLeft      =   1890
   ClientTop       =   2265
   ClientWidth     =   8670
   Icon            =   "frmDyeAux.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   8670
   Begin Threed.SSPanel pnlMsg 
      Height          =   510
      Left            =   4050
      TabIndex        =   26
      Top             =   5070
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
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   4905
      Left            =   15
      TabIndex        =   25
      Top             =   975
      Width           =   3840
      _cx             =   6773
      _cy             =   8652
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
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   2085
      TabIndex        =   24
      Top             =   5970
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&D)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlEdit 
      Height          =   1905
      Left            =   3990
      TabIndex        =   9
      Top             =   1065
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   3360
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtBox 
         Height          =   300
         Index           =   4
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1515
         Width           =   3300
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   300
         Index           =   3
         Left            =   1170
         TabIndex        =   13
         Top             =   1155
         Width           =   1200
      End
      Begin VB.TextBox txtBox 
         Height          =   300
         Index           =   2
         Left            =   1170
         MaxLength       =   3
         TabIndex        =   12
         Top             =   795
         Width           =   1185
      End
      Begin VB.TextBox txtBox 
         Height          =   300
         Index           =   0
         Left            =   1170
         MaxLength       =   3
         TabIndex        =   10
         Top             =   75
         Width           =   570
      End
      Begin VB.TextBox txtBox 
         Height          =   300
         Index           =   1
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   11
         Top             =   435
         Width           =   3300
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   17
         Top             =   75
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "코드"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   20
         Top             =   435
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "조제명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   105
         TabIndex        =   21
         Top             =   795
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "단위"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   105
         TabIndex        =   22
         Top             =   1155
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "단가"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   6
         Left            =   105
         TabIndex        =   23
         Top             =   1515
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "비고사항"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   6960
      TabIndex        =   16
      Top             =   5955
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin TabDlg.SSTab tabCode 
      Height          =   5325
      Left            =   3900
      TabIndex        =   15
      Top             =   975
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   9393
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   750
      TabCaption(0)   =   "  조제 관리  "
      TabPicture(0)   =   "frmDyeAux.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "  염료 관리  "
      TabPicture(1)   =   "frmDyeAux.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin Threed.SSPanel pnlBoard 
      Height          =   930
      Left            =   15
      TabIndex        =   18
      Top             =   15
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   1640
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optSize 
         Caption         =   "요약"
         Height          =   330
         Index           =   0
         Left            =   1440
         Style           =   1  '그래픽
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   495
         Width           =   645
      End
      Begin VB.OptionButton optSize 
         Caption         =   "상세"
         Height          =   330
         Index           =   1
         Left            =   1440
         Style           =   1  '그래픽
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   105
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   4590
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   4
         ToolTipText     =   "자료 저장"
         Top             =   75
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   6180
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   6
         ToolTipText     =   "자료 추가"
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   7770
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   8
         ToolTipText     =   "자료 삭제"
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   6975
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   7
         ToolTipText     =   "자료 수정"
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   5385
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   5
         ToolTipText     =   "자료 취소"
         Top             =   75
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.ComboBox cboCode 
         Height          =   300
         Left            =   120
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   495
         Width           =   1245
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "코드종류"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
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
      Left            =   180
      TabIndex        =   19
      Top             =   6015
      Width           =   75
   End
End
Attribute VB_Name = "frmDyeAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sFlag        As String * 1

Private Sub cmdExcel_Click()
    If grdData.Rows = grdData.FixedRows Then
        MsgBox LoadResString(203), vbInformation
        Exit Sub
    End If
    
    Call MakeExcelGrid(grdData)
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 8790, 7125

    Call SetOperate(Me)
    Call InitGrid

    ' 콤보 설정
    With cboCode
        .AddItem "조제 관리"
        .AddItem "염료 관리"
        .ListIndex = 0
    End With
    tabCode.Tab = 0
End Sub

Private Sub cboCode_Click()
    tabCode.Tab = cboCode.ListIndex
    Call FillGrid
End Sub

Private Sub grdData_AfterSort(ByVal Col As Long, Order As Integer)
    Call ShowData
End Sub

Private Sub optSize_Click(Index As Integer)
    With grdData
        If optSize(0).Value Then    ' 요약
            .Width = 8655

            tabCode.Visible = False
            pnlEdit.Visible = False
            
            .ColHidden(3) = False
            .ColHidden(4) = False
            .ColHidden(5) = False
        Else                        ' 상세
            .Width = 3870
            
            tabCode.Visible = True
            pnlEdit.Visible = True
            
            .ColHidden(3) = True
            .ColHidden(4) = True
            .ColHidden(5) = True
        End If
    End With
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Select Case Index
        Case ID_ADDNEW
            optSize(1).Value = True
            Call ChangeMode(Me, False)
            Call ClearData
            pnlMsg.Caption = LoadResString(121)
            cboCode.Enabled = False
            tabCode.Enabled = False
            m_sFlag = ID_ADDNEW

            txtBox(0).SetFocus
        Case ID_UPDATE
            optSize(1).Value = True
            Call ChangeMode(Me, False)
            pnlMsg.Caption = LoadResString(122)
            cboCode.Enabled = False
            tabCode.Enabled = False
            m_sFlag = ID_UPDATE

            txtBox(0).Locked = True
            txtBox(1).SetFocus
        Case ID_DELETE
            If grdData.Rows = grdData.FixedRows Then Exit Sub

            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") <> vbYes Then Exit Sub

            If DeleteData() Then Call FillGrid
        Case ID_SAVE
            If SaveData() Then
                cboCode.Enabled = True
                tabCode.Enabled = True
                Call FillGrid
                Call ChangeMode(Me, True)
            End If
        Case ID_CANCEL
            m_sFlag = ""
            If grdData.Rows > grdData.FixedRows Then
                Call ShowData
            Else
                Call ClearData
            End If
            txtBox(0).Locked = False
            cboCode.Enabled = True
            tabCode.Enabled = True
            Call ChangeMode(Me, True)
            grdData.SetFocus
    End Select
End Sub


Private Sub grdData_RowColChange()
    Call ShowData
End Sub

Private Sub txtBox_GotFocus(Index As Integer)
    If Index = 3 Then
        
        txtBox(Index) = CStr(CLng(txtBox(Index)))

    End If
    
    Call GotFocusText(txtBox(Index))
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtBox_LostFocus(Index As Integer)
    If Index = 3 Then
        txtBox(3) = SetCurrency(txtBox(3))
    End If
    
End Sub



Private Sub tabCode_Click(PreviousTab As Integer)
    cboCode.ListIndex = tabCode.Tab
    Select Case tabCode.Tab
        Case 0
            pnlCaption(0).Caption = "조제명"
        Case 1
            pnlCaption(0).Caption = "염료명"
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    With grdData
        .Cols = 6
        Call SetVSFlexGrid(grdData)

        .Redraw = False

        .TextArray(0) = ""
        .TextArray(1) = "코드":     .ColWidth(1) = 555:         .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "염조제명": .ColWidth(2) = 2500:        .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "단위":     .ColHidden(3) = 330:         .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "단가":     .ColWidth(4) = 690:         .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "비고사항": .ColWidth(5) = 1845:        .ColAlignment(5) = flexAlignLeftCenter

        .ColHidden(3) = True
        .ColHidden(4) = True
        .ColHidden(5) = True
        
        .Redraw = True
    End With
End Sub

Private Sub FillGrid()
    Dim oDyeAux As PlusLib2.CDyeAux
    Dim rs As ADODB.Recordset
    Dim i%, iNowRow%

    On Error GoTo ErrHandler

    Set oDyeAux = New PlusLib2.CDyeAux
    oDyeAux.Connection = g_adoCon

    Set rs = oDyeAux.GetDyeAux(CStr(tabCode.Tab))
    Set oDyeAux = Nothing

    With grdData
        .Redraw = False

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows

        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & CStr(rs!DyeAuxID) & vbTab & CStr(rs!DyeAux) & vbTab & _
                CheckNull(rs!Unit) & vbTab & SetCurrency(rs!UnitPrice) & vbTab & _
                CheckNull(rs!Remark)

            rs.MoveNext
        Next i

        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
             Call ShowData
        Else
            grdData.HighLight = flexHighlightNever
            Call ClearData
        End If

        lblCount = LoadResString(250) & CStr(.Rows - .FixedRows) & " 건"

        .Redraw = True
    End With

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oDyeAux = Nothing
    
    Call ErrorBox(Err.Number, "DyeAux", Err.Description)
End Sub

Private Sub ShowData()
    With grdData
    
        txtBox(0) = Right(.TextMatrix(.Row, 1), 3)
        txtBox(1) = .TextMatrix(.Row, 2)
        txtBox(2) = .TextMatrix(.Row, 3)
        txtBox(3) = .TextMatrix(.Row, 4)
        txtBox(4) = .TextMatrix(.Row, 5)
        
    End With
End Sub

Private Sub ClearData()
    txtBox(0) = ""
    txtBox(1) = ""
    txtBox(2) = ""
    txtBox(3) = "0"
    txtBox(4) = ""
End Sub

Private Function CheckData() As Boolean
    CheckData = False

    If Len(txtBox(0)) = 0 Then
        MsgBox "'코드'를 입력하십시오.", vbInformation
        txtBox(0).SetFocus
        Exit Function
    End If
    If Len(txtBox(1)) = 0 Then
        MsgBox "'" & IIf(tabCode.Tab = 0, "조제", "염료") & "명'을 입력하십시오.", vbInformation
        txtBox(1).SetFocus
        Exit Function
    End If
    If Len(txtBox(2)) = 0 Then
        MsgBox "'단위'를 입력하십시오.", vbInformation
        txtBox(2).SetFocus
        Exit Function
    End If

    CheckData = True
End Function

Private Function SaveData() As Boolean
    Dim oDyeAux As PlusLib2.CDyeAux
    Dim tDA     As PlusLib2.TDyeAux
    Dim rs As ADODB.Recordset
    Dim sDyeAuxID$

    On Error GoTo ErrHandler

    With tDA
        .DyeAuxID = Format(txtBox(0), "000")
        .DyeAux = txtBox(1)
        .Unit = txtBox(2)
        .UnitCost = CLng(txtBox(3))
        .Remark = txtBox(4)
        .nKind = CStr(tabCode.Tab)
    End With

    If m_sFlag = ID_ADDNEW Then
        Set oDyeAux = New PlusLib2.CDyeAux
        oDyeAux.Connection = g_adoCon
        oDyeAux.UserName = g_sUserName
        
        sDyeAuxID = CStr(tDA.nKind) & tDA.DyeAuxID
        
        Set rs = oDyeAux.GetDyeAuxOne(sDyeAuxID)
    
        Set oDyeAux = Nothing
    
        If Not rs.EOF Then
            MsgBox "동일한 코드번호를 가진 염조제가 있습니다" & vbCrLf & "염조제 코드를 다시입력하십시오"
            
            rs.Close
            Set rs = Nothing
            Exit Function
        End If
        
        rs.Close
        Set rs = Nothing
    End If
    
    Set oDyeAux = New PlusLib2.CDyeAux
    oDyeAux.Connection = g_adoCon
    oDyeAux.UserName = g_sUserName
    

    If m_sFlag = ID_ADDNEW Then
        SaveData = oDyeAux.AddNewDyeAux(tDA)
    Else
        SaveData = oDyeAux.UpdateDyeAux(tDA)
    End If

    Set oDyeAux = Nothing

    Exit Function

ErrHandler:
    Set rs = Nothing
    Set oDyeAux = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Function

Private Function DeleteData() As Boolean
    Dim oDyeAux As PlusLib2.CDyeAux
    Dim sDyeAuxID$

    On Error GoTo ErrHandler

    DeleteData = False

    Set oDyeAux = New PlusLib2.CDyeAux
    oDyeAux.Connection = g_adoCon
    oDyeAux.UserName = g_sUserName

    sDyeAuxID = CStr(tabCode.Tab) & txtBox(0)
    DeleteData = oDyeAux.DeleteDyeAux(sDyeAuxID)

    Set oDyeAux = Nothing

    Exit Function

ErrHandler:
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Set oDyeAux = Nothing
End Function

