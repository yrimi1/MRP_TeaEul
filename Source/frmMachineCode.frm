VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMachineCode 
   BorderStyle     =   1  '단일 고정
   Caption         =   "공정/기계 관리"
   ClientHeight    =   7275
   ClientLeft      =   2820
   ClientTop       =   1935
   ClientWidth     =   10725
   Icon            =   "frmMachineCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10725
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   390
      Top             =   6525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   690
      Left            =   9000
      TabIndex        =   1
      Top             =   6480
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlEdit 
      Height          =   5325
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   990
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   9393
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   5175
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   90
         Width           =   3330
         _cx             =   5874
         _cy             =   9128
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
      Height          =   5325
      Index           =   1
      Left            =   3555
      TabIndex        =   3
      Top             =   990
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   9393
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   5205
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   3495
         _cx             =   6165
         _cy             =   9181
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
   Begin Threed.SSPanel pnlBoard 
      Height          =   915
      Left            =   3555
      TabIndex        =   5
      Top             =   30
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   1614
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   3090
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   10
         ToolTipText     =   "자료 저장"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   4680
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   9
         ToolTipText     =   "자료 추가"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   6270
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   8
         ToolTipText     =   "자료 삭제"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   5475
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   7
         ToolTipText     =   "자료 수정"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   3885
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   6
         ToolTipText     =   "자료 취소"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   900
      Left            =   30
      TabIndex        =   11
      Top             =   45
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   1588
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   75
         TabIndex        =   12
         Top             =   465
         Width           =   2025
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   90
         TabIndex        =   13
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
         TabIndex        =   14
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
   Begin Threed.SSPanel pnlEdit 
      Height          =   5295
      Index           =   2
      Left            =   7245
      TabIndex        =   15
      Top             =   1005
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   9340
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   0
         Left            =   1155
         TabIndex        =   19
         Top             =   450
         Width           =   2220
         _ExtentX        =   3916
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
         Left            =   1155
         TabIndex        =   18
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
         Locked          =   -1  'True
         BackColor       =   12648384
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "코      드"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   450
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "기 계 명"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MRPPlus2.WizText txtName 
         Height          =   300
         Index           =   1
         Left            =   1155
         TabIndex        =   20
         Top             =   810
         Width           =   2220
         _ExtentX        =   3916
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
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   810
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "호     기"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   510
         Left            =   90
         TabIndex        =   22
         Top             =   4320
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   900
         _Version        =   196609
         BackColor       =   65535
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMachineCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum GroupIndex
    GI_Large = 0
    GI_Middle = 1
End Enum

Private Const LIMIT_WIDTH1 = 2860
Private Const LIMIT_WIDTH2 = 2410
Private Const LIMIT_WIDTH3 = 3000
Private Const LIMIT_ROW = 18

Dim m_bSkip As Boolean
Dim m_sFlag As String * 1


Private Sub cmdOperate_Click(Index As Integer)
    Dim bResult As Boolean
        
    On Error GoTo ErrHandler
    '---------------------------------------------------------------------------
    Select Case Index   '[1] 추가
        Case ID_ADDNEW
            m_sFlag = ID_ADDNEW
            Call ChangeMode(Me, False)
            Call ClearData
            pnlMsg.Caption = LoadResString(302)
            
            txtName(0).SetFocus
    '---------------------------------------------------------------------------
        Case ID_UPDATE '[2] 수정
            m_sFlag = ID_UPDATE
            Call ChangeMode(Me, False)
            
            pnlMsg.Caption = LoadResString(303)
            txtCode.Locked = True
            txtName(0).SetFocus
    '---------------------------------------------------------------------------
        Case ID_DELETE '[3] 삭제
            If grdData(0).Rows = grdData(0).FixedRows Then Exit Sub
    
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                m_sFlag = ID_DELETE
                
                If SaveData() Then Call FillGrid(GI_Large)
            End If
    '---------------------------------------------------------------------------
        Case ID_SAVE  '[4] 저장
            'If CheckData() = False Then Exit Sub
            If SaveData() Then
                Call FillGrid(GI_Large)
                Call ChangeMode(Me, True)
                
                m_sFlag = ""
                txtCode.Locked = False
            End If
            
            grdData(1).SetFocus
    '---------------------------------------------------------------------------
        Case ID_CANCEL '[5] 취소
            m_sFlag = ""
            If grdData(0).Rows > 1 Then
                Call ShowData
            Else
                Call ClearData
            End If
            Call ChangeMode(Me, True)
            txtCode.Locked = False
            grdData(1).SetFocus
    End Select
    
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "Code.cmdOperate_Click", Err.Description)
End Sub


Private Sub ClearData()
    txtCode = ""
    txtName(0) = ""
    txtName(1) = ""

End Sub


Private Sub ShowData()
    
    If grdData(1).Rows = grdData(1).FixedRows Then
        Call ClearData
        Exit Sub
    End If
    
    With grdData(1)
        txtCode = .TextMatrix(.Row, 1)
        txtName(0) = .TextMatrix(.Row, 2)
        txtName(1) = .TextMatrix(.Row, 3)
    End With
End Sub


Private Function SaveData() As Boolean
    Dim NewCode As PlusLib2.TMachine
    Dim oProcess As PlusLib2.CProcess

    On Error GoTo ErrHandler
    
    Set oProcess = New PlusLib2.CProcess
    oProcess.Connection = g_adoCon
    oProcess.UserName = g_sUserName
    
    NewCode.sProcessID = grdData(0).TextMatrix(grdData(0).Row, 1)
    NewCode.sMachineID = Format(txtCode, "00")
    NewCode.sMachine = txtName(0)
    NewCode.sMachineNo = txtName(1)
    
    
    If m_sFlag = ID_ADDNEW Then
        SaveData = oProcess.AddNewMachine(NewCode)
    ElseIf m_sFlag = ID_UPDATE Then
        SaveData = oProcess.UpdateMachine(NewCode)
    ElseIf m_sFlag = ID_DELETE Then
        SaveData = oProcess.DeleteMachine(NewCode.sProcessID, NewCode.sMachineID)
    End If
    
    Set oProcess = Nothing
    Exit Function
    
ErrHandler:
    Set oProcess = Nothing

    Call ErrorBox(Err.Number, "frmMachineCode.SaveData", Err.Description)
End Function


Private Sub Form_Load()
        
    Me.Move 0, 0, 10820, 7660
 
    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
    cmdExit.MousePointer = ssCustom
    cmdExit.MouseIcon = LoadResPicture("POINTER", vbResCursor)

    Call InitGrid
    
    Call FillGrid(GI_Large)
    

    m_bSkip = False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMachineCode = Nothing
    
End Sub


Private Sub grdData_RowColChange(Index As Integer)
    If Index = GI_Large Then Call FillGrid(GI_Middle)
    
    If Index = 1 Then Call ShowData
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub InitGrid()
    
    ' 공정관리 Grid
    Call SetVSFlexGrid(grdData(0))
    With grdData(0)
        .Redraw = False
        .Cols = 3
        
        .TextArray(0) = "":             .ColWidth(0) = 330
        .TextArray(1) = "코드":         .ColWidth(1) = 450
        .TextArray(2) = "공정명":       .ColWidth(2) = LIMIT_WIDTH1
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .Redraw = True
    End With
    ' 기계관리 Grid
    
    Call SetVSFlexGrid(grdData(1))
    With grdData(1)
        .Redraw = False
        .Cols = 4
        
        .TextArray(0) = "":                         .ColWidth(0) = 330
        .TextArray(1) = "코드":                     .ColWidth(1) = 450
        .TextArray(2) = "기계명":                   .ColWidth(2) = 2000
        .TextArray(3) = "기계" & vbCrLf & "호기":   .ColWidth(3) = 450
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .Redraw = True
    End With

End Sub



Private Function FillGrid(ByVal Index As GroupIndex) As Boolean
    Dim oProcess As PlusLib2.CProcess
    Dim rs As ADODB.Recordset
    Dim iLoop%
    Dim sKey As String

    On Error GoTo ErrHandler
    
    Set oProcess = New PlusLib2.CProcess
    oProcess.Connection = g_adoCon
    If Index = GI_Large Then
        Set rs = oProcess.GetProcess()
        
    ElseIf Index = GI_Middle Then
        With grdData(0)
            sKey = .TextMatrix(.Row, 1)
        End With
        If Len(sKey) = 0 Then
            grdData(1).Rows = 1
            grdData(1).HighLight = flexHighlightNever
            
            Exit Function
        Else
            Set rs = oProcess.GetMachine(sKey)
        End If
    End If
    Set oProcess = Nothing
    
    Dim lNowRow As Long
    
    With grdData(Index)
        If .Rows > .FixedRows Then
            lNowRow = .Row
            .Rows = 1
        Else
            lNowRow = 1
        End If
        
        If Index = GI_Large Then
            Do Until rs.EOF
                .AddItem CStr(.Rows) & vbTab & rs!PROCESSID & vbTab & CheckNull(rs!Process)
    
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                .AddItem CStr(.Rows) & vbTab & rs!MachineID & vbTab & rs!Machine & vbTab & CheckNull(rs!MachineNo)
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
        
        
        If .Rows = .FixedRows Then
            .HighLight = flexHighlightNever
            
            If Index = GI_Large Then
                grdData(1).Rows = 1
                grdData(1).HighLight = flexHighlightNever
            Else
                Call ClearData
            
            End If
        Else
            .HighLight = flexHighlightAlways
            
            If .Rows > lNowRow Then
                .Row = lNowRow
            Else
                .Row = .Rows - 1
            End If
            
            .Col = .FixedCols
            .ColSel = .Cols - 1
            If Index = GI_Large Then Call FillGrid(GI_Middle)
        End If
    End With
    Exit Function

ErrHandler:
    MsgBox CStr(Err.Number) & Err.Description, vbCritical
    Err.Clear
    Set rs = Nothing
    Set oProcess = Nothing
End Function

