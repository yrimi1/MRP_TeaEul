VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmDyeAuxGroup 
   Caption         =   "염조제 그룹 관리"
   ClientHeight    =   7305
   ClientLeft      =   1575
   ClientTop       =   1500
   ClientWidth     =   8670
   Icon            =   "frmDyeAuxGroup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   8670
   Begin Threed.SSPanel pnlMsg 
      Height          =   510
      Left            =   4305
      TabIndex        =   7
      Top             =   6720
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
      Left            =   6990
      TabIndex        =   6
      Top             =   6570
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlBoard 
      Height          =   915
      Index           =   0
      Left            =   4080
      TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   0
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
         TabIndex        =   3
         ToolTipText     =   "자료 저장"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   900
      Left            =   30
      TabIndex        =   8
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
         TabIndex        =   9
         Top             =   465
         Width           =   2025
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   90
         TabIndex        =   10
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
         TabIndex        =   11
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
      Height          =   5460
      Left            =   30
      TabIndex        =   13
      Top             =   1005
      Width           =   3975
      _cx             =   7011
      _cy             =   9631
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
   Begin Threed.SSPanel pnlDyeAux 
      Height          =   4185
      Left            =   4095
      TabIndex        =   14
      Top             =   2280
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   7382
      _Version        =   196609
      Enabled         =   0   'False
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtTemp 
         Height          =   300
         Left            =   210
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   195
         Visible         =   0   'False
         Width           =   270
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   540
         Left            =   3255
         TabIndex        =   16
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   953
         _Version        =   196609
         Caption         =   "염조제 삭제"
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   540
         Left            =   2040
         TabIndex        =   17
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   953
         _Version        =   196609
         Caption         =   "염조제추가"
      End
      Begin VSFlex7LCtl.VSFlexGrid grdDyeAux 
         Height          =   3450
         Left            =   75
         TabIndex        =   18
         Top             =   660
         Width           =   4350
         _cx             =   7673
         _cy             =   6085
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
   Begin Threed.SSPanel pnlEdit 
      Height          =   1200
      Left            =   4095
      TabIndex        =   19
      Top             =   1005
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   2117
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtName 
         Height          =   330
         Left            =   1170
         TabIndex        =   21
         Top             =   435
         Width           =   3240
      End
      Begin VB.TextBox txtPerson 
         Height          =   330
         Left            =   1170
         TabIndex        =   20
         Top             =   810
         Width           =   2865
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   315
         Index           =   0
         Left            =   4080
         TabIndex        =   22
         Top             =   810
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   196609
         PictureFrames   =   1
         Picture         =   "frmDyeAuxGroup.frx":000C
         ButtonStyle     =   3
      End
      Begin MRPPlus2.WizText txtCode 
         Height          =   300
         Left            =   1185
         TabIndex        =   23
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
         MaxLength       =   4
         BackColor       =   12648384
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   90
         TabIndex        =   24
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
         TabIndex        =   25
         Top             =   450
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "그 룹 명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   90
         TabIndex        =   26
         Top             =   810
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "작 성 자"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
      Left            =   60
      TabIndex        =   12
      Top             =   6930
      Width           =   945
   End
End
Attribute VB_Name = "frmDyeAuxGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIMIT_WIDTH = 3140
Private Const LIMIT_ROW = 16

Private m_sFlag        As String * 1
Private m_bSortForward As Boolean


Private Sub cmdAdd_Click()
    With grdDyeAux
        .SetFocus
        .AddItem CStr(.Rows)

        .Cell(flexcpPicture, .Rows - 1, 3) = LoadResPicture("B_FIND", vbResBitmap)
        .Cell(flexcpPictureAlignment, .Rows - 1, 3) = flexPicAlignCenterCenter
        
        .Cell(flexcpPicture, .Rows - 1, 4) = LoadResPicture("B_FIND", vbResBitmap)
        .Cell(flexcpPictureAlignment, .Rows - 1, 4) = flexPicAlignCenterCenter

        .Select .Rows - 1, 1
    End With
End Sub


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


Private Sub cmdDel_Click()
    With grdDyeAux
        If .Rows = .FixedRows Or .Row < .FixedRows Or .Row >= .Rows Then Exit Sub

        .RemoveItem .Row
    End With

End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub ClearData()

    txtCode = ""
    txtName = ""
    txtPerson = ""

    grdDyeAux.Rows = grdDyeAux.FixedRows

End Sub


Private Sub cmdFind_Click(Index As Integer)
    Call ReturnCode(LG_PERSON, , False, txtPerson)
End Sub


Private Sub cmdOperate_Click(Index As Integer)
    Dim bResult As Boolean
        
    On Error GoTo ErrHandler
    '---------------------------------------------------------------------------
    Select Case Index   '[1] 추가
        Case ID_ADDNEW
            m_sFlag = ID_ADDNEW
            Call ChangeMode(Me, False)
            
            Call ClearData
            txtName.SetFocus
            
            pnlDyeAux.Enabled = True
            
            pnlMsg.Caption = LoadResString(302)
            
    '---------------------------------------------------------------------------
        Case ID_UPDATE '[2] 수정
        
            If grdData.Rows = grdData.FixedRows Then Exit Sub
            
            pnlEdit.Enabled = True
            pnlDyeAux.Enabled = True
            
            txtCode.Locked = True
            txtName.SetFocus
            
            
            m_sFlag = ID_UPDATE
            Call ChangeMode(Me, False)
            
            pnlMsg.Caption = LoadResString(303)
            
    '---------------------------------------------------------------------------
        Case ID_DELETE '[3] 삭제
        
            If grdData.Rows = grdData.FixedRows Then Exit Sub
    
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                m_sFlag = ID_DELETE
                
                If SaveData() Then Call FillGridGroup
            End If
    '---------------------------------------------------------------------------
        Case ID_SAVE  '[4] 저장
'            m_sFlag = ID_SAVE
            If CheckData() = False Then Exit Sub
            If SaveData() Then
                Call FillGridGroup
                Call ChangeMode(Me, True)
                
                m_sFlag = ""
                                
                txtCode.Locked = False
                
                pnlDyeAux.Enabled = False
                pnlEdit.Enabled = False
            
            End If
            grdData.SetFocus
    '---------------------------------------------------------------------------
        Case ID_CANCEL '[5] 취소
            m_sFlag = ""
            If grdData.Rows > 1 Then
                Call ShowData
            End If
            
            Call ChangeMode(Me, True)
            
            txtCode.Locked = False
                
            pnlDyeAux.Enabled = False
            pnlEdit.Enabled = False
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
        txtPerson = .TextMatrix(.Row, 3)
        txtPerson.Tag = .TextMatrix(.Row, 4)
    End With
    
    Call FillGridDyeAux

End Sub

Private Sub Form_Activate()
'    txtSearch.SetFocus
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 8790, 7710
    
    Call SetOperate(Me)
    
    Call InitGrid
    Call FillGridGroup
    
    cmdAll.Picture = LoadResPicture("ALL", vbResIcon)
    lblCount.Caption = LoadResString(250)
End Sub

Private Sub InitGrid()
    Call SetVSFlexGrid(grdData)
    With grdData
        .Redraw = False
        .Rows = 1
        .Cols = 5
        
        .TextArray(0) = ""
        .TextArray(1) = "코드":         .ColWidth(1) = 450:     .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "염조제 그룹":  .ColWidth(2) = 1200:    .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "작성자":       .ColWidth(3) = 0
        .TextArray(4) = "작성자ID":     .ColWidth(4) = 0
        
        .Redraw = True
    End With
    
    Call SetVSFlexGrid(grdDyeAux)
    With grdDyeAux
        .Redraw = False
        .Rows = 1
        .Cols = 5
        
        .TextArray(0) = " "
        .TextArray(1) = "염조제ID":       .ColWidth(1) = 900:   .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "염조제명":       .ColWidth(2) = 1500:  .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "염료":           .ColWidth(3) = 600:   .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "조제":           .ColWidth(4) = 600:   .ColAlignment(4) = flexAlignCenterCenter
      
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True

        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusHeavy

        .Redraw = flexRDDirect

    End With
    
End Sub


Private Sub grdData_RowColChange()
    Call ShowData
End Sub


Private Sub grdDyeAux_Click()
    With grdDyeAux
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub
        
       ' If .MouseCol > 3 Or .MouseCol > 4 Then Exit Sub

        Dim iRow%
        iRow = .MouseRow

        txtTemp = ""
        If .MouseCol = 3 Then
            If ReturnCode(LG_DYE, , , txtTemp) Then
                .TextMatrix(.Row, 1) = txtTemp.Tag
                .TextMatrix(.Row, 2) = txtTemp
            Else
                .TextMatrix(.Row, 1) = ""
                .TextMatrix(.Row, 2) = ""
            End If
        ElseIf .MouseCol = 4 Then
            If ReturnCode(LG_AUX, , , txtTemp) Then
                .TextMatrix(.Row, 1) = txtTemp.Tag
                .TextMatrix(.Row, 2) = txtTemp
            Else
                .TextMatrix(.Row, 1) = ""
                .TextMatrix(.Row, 2) = ""
            End If
        End If
    End With

End Sub


Private Sub tabform_Click(PreviousTab As Integer)
    Dim sMenuID As String
    
    Call FillGridGroup
        
    txtSearch.SetFocus
End Sub



Private Sub FillGridGroup()
    Dim oDyeAux As PlusLib2.CDyeAux
    Dim rs As ADODB.Recordset
    Dim lNowRow&
    
    On Error GoTo ErrHandler
        
    Set oDyeAux = New PlusLib2.CDyeAux
    oDyeAux.Connection = g_adoCon
        
    Set rs = oDyeAux.GetDyeAuxGroup
    Set oDyeAux = Nothing
    
    If rs.RecordCount = 0 Then
        grdData.Rows = grdData.FixedRows
        grdData.HighLight = flexHighlightNever
        lblCount.Caption = LoadResString(250)
        
        rs.Close
        Set rs = Nothing
        
        Call ClearData
        
        Exit Sub
    End If
    
    With grdData
        .Redraw = False
        
        lNowRow = IIf(.Row > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs!GroupID & vbTab & rs!GroupName & vbTab & rs!Name & vbTab & rs!PersonID
            rs.MoveNext
        Loop
            
        lblCount.Caption = LoadResString(250) & grdData.Rows - 1 & " 건"
        
        rs.Close
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
           .HighLight = flexHighlightAlways
           .Row = IIf(.Rows > lNowRow, lNowRow, .Rows - 1)
           .TopRow = lNowRow
           
           .Col = .FixedCols
           .ColSel = .Cols - 1

        End If
        .Redraw = True
    End With
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oDyeAux = Nothing
    
    Call ErrorBox(Err.Number, "Code.FillGrid", Err.Description)
End Sub


Private Sub FillGridDyeAux()
    Dim oDyeAux As PlusLib2.CDyeAux
    Dim rs As ADODB.Recordset
    Dim lNowRow&
    Dim sGroupID$
    
    On Error GoTo ErrHandler
        
    With grdData
        sGroupID = .TextMatrix(.Row, 1)
    End With
        
    Set oDyeAux = New PlusLib2.CDyeAux
    oDyeAux.Connection = g_adoCon
        
    Set rs = oDyeAux.GetDyeAuxGroupSub(sGroupID)
    Set oDyeAux = Nothing
    
    If rs.RecordCount = 0 Then
        grdDyeAux.Rows = grdDyeAux.FixedRows
        grdDyeAux.HighLight = flexHighlightNever
                
        rs.Close
        Set rs = Nothing
        
        Call ClearData

        Exit Sub
    End If
    
    With grdDyeAux
        .Redraw = False
        
        lNowRow = IIf(.Row > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs!DyeAuxID & vbTab & rs!DyeAux
            
            .Cell(flexcpPicture, .Rows - 1, 3) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, .Rows - 1, 3) = flexPicAlignCenterCenter
            
            .Cell(flexcpPicture, .Rows - 1, 4) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, .Rows - 1, 4) = flexPicAlignCenterCenter
            
            rs.MoveNext
        Loop
        
        lblCount.Caption = LoadResString(250) & grdData.Rows - 1 & " 건"
        
        rs.Close
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
           .HighLight = flexHighlightAlways
           .Row = IIf(.Rows > lNowRow, lNowRow, .Rows - 1)
           .TopRow = lNowRow
           
           .Col = .FixedCols
           .ColSel = .Cols - 1
           
        End If
        .Redraw = True
    End With
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oDyeAux = Nothing
    
    Call ErrorBox(Err.Number, "Code.FillGrid", Err.Description)
End Sub



Private Function CheckData() As Boolean
    Dim i%, iCntDyeAux%
    
    CheckData = True
    If m_sFlag = ID_ADDNEW Or m_sFlag = ID_UPDATE Then
        If Len(txtName) = 0 Then
            MsgBox LoadResString(115), vbInformation
            txtName.SetFocus
            CheckData = False
            Exit Function
        End If
    ElseIf m_sFlag = ID_SAVE Then
        If Len(Trim(txtName)) = 0 Then
            MsgBox "그룹명이 명시되어 있지 않습니다", vbInformation, "그룹명"
            CheckData = False
            Exit Function
        End If
        With grdDyeAux
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    iCntDyeAux = iCntDyeAux + 1
                End If
            Next i
        End With
        
        If grdDyeAux.Rows = grdDyeAux.FixedRows Or iCntDyeAux = 0 Then
            MsgBox "추가할 염조제 항목이 없습니다", vbInformation, "염조제 미등록"
            CheckData = False
            Exit Function
        End If
    End If
End Function

Private Function SaveData() As Boolean
    Dim NewGroup As PlusLib2.TDyeAuxGroup
    Dim NewGroupSub() As PlusLib2.TDyeAuxGroupSub
    Dim oDyeAux As PlusLib2.CDyeAux
    Dim nCnt%, i%, idx%
    Dim sGroupID$, nSeq%
    
    
    On Error GoTo ErrHandler
    
    Set oDyeAux = New PlusLib2.CDyeAux
    oDyeAux.Connection = g_adoCon
    oDyeAux.UserName = g_sUserName
    

    NewGroup.GroupID = Format(txtCode, "0000")
    NewGroup.GroupName = Trim(txtName)
    NewGroup.PersonID = txtPerson.Tag
    
    If m_sFlag = ID_DELETE Then
        SaveData = oDyeAux.DeleteDyeAuxGroup(NewGroup.GroupID)
        
    Else
    
'        nCnt = -1
        With grdDyeAux
            If .Rows > .FixedRows Then
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, 1)) <> "" Then
                        nCnt = nCnt + 1
                    End If
                Next i
                
                If nCnt > 0 Then
                    ReDim NewGroupSub(nCnt)
        
                    For i = .FixedRows To .Rows - 1
                        If Trim(.TextMatrix(i, 1)) <> "" Then
                            NewGroupSub(idx).GroupID = Format(txtCode, "0000")
                            NewGroupSub(idx).Seq = idx + 1
                            NewGroupSub(idx).DyeAuxID = .TextMatrix(i, 1)
                            idx = idx + 1
                        End If
                    Next i
                Else
                    Exit Function
                End If
            End If
            
        End With
    
        If m_sFlag = ID_ADDNEW Then
            SaveData = oDyeAux.AddNewDyeAuxGroup(NewGroup, NewGroupSub, nCnt)

        ElseIf m_sFlag = ID_UPDATE Then
            SaveData = oDyeAux.UpdateDyeAuxGroup(NewGroup, NewGroupSub, nCnt)
            
        End If
    End If
    

    
        
    Set oDyeAux = Nothing
    Exit Function
ErrHandler:
    Set oDyeAux = Nothing
    Call ChangeMode(Me, True)
    Call ErrorBox(Err.Number, "frmDyeAuxGroup.SaveData", Err.Description)
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


    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "txtSearch.Change", Err.Description)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call MoveFocus(KeyCode)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        grdData.SetFocus
    End If
End Sub
