VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmPatternCode 
   Caption         =   "공정패턴 관리"
   ClientHeight    =   7440
   ClientLeft      =   3000
   ClientTop       =   1980
   ClientWidth     =   11865
   Icon            =   "frmPatternCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   11865
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   795
      Index           =   1
      Left            =   3360
      TabIndex        =   25
      Top             =   60
      Width           =   8475
      _cx             =   14949
      _cy             =   1402
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
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   6555
      Index           =   0
      Left            =   15
      TabIndex        =   24
      Top             =   60
      Width           =   3285
      _cx             =   5794
      _cy             =   11562
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
   Begin Threed.SSPanel pnlBoard 
      Height          =   5730
      Left            =   3360
      TabIndex        =   18
      Top             =   900
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   10107
      _Version        =   196609
      Caption         =   "SSPanel1"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   780
         Index           =   4
         Left            =   5235
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   13
         ToolTipText     =   "자료 취소"
         Top             =   75
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   6825
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   15
         ToolTipText     =   "자료 수정"
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   7620
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   16
         ToolTipText     =   "자료 삭제"
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   6030
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   14
         ToolTipText     =   "자료 추가"
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   780
         Index           =   3
         Left            =   4440
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   12
         ToolTipText     =   "자료 저장"
         Top             =   75
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSPanel pnlEdit 
         Height          =   4725
         Left            =   60
         TabIndex        =   0
         Top             =   930
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   8334
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboWork 
            Height          =   300
            Left            =   1305
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   390
            Width           =   1050
         End
         Begin Threed.SSFrame fraProcess 
            Height          =   3600
            Left            =   45
            TabIndex        =   7
            Top             =   1065
            Width           =   5145
            _ExtentX        =   9075
            _ExtentY        =   6350
            _Version        =   196609
            Begin VSFlex7LCtl.VSFlexGrid grdPattern 
               Height          =   3075
               Left            =   2985
               TabIndex        =   23
               Top             =   435
               Width           =   2070
               _cx             =   3651
               _cy             =   5424
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
            Begin VSFlex7LCtl.VSFlexGrid grdProcess 
               Height          =   3075
               Left            =   90
               TabIndex        =   22
               Top             =   435
               Width           =   2040
               _cx             =   3598
               _cy             =   5424
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
            Begin VB.CommandButton cmdMove 
               Caption         =   ">>"
               Height          =   555
               Index           =   0
               Left            =   2295
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   1350
               Width           =   555
            End
            Begin VB.CommandButton cmdMove 
               Caption         =   "<<"
               Height          =   555
               Index           =   1
               Left            =   2295
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   2010
               Width           =   555
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   0
               Left            =   90
               TabIndex        =   8
               Top             =   90
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "전체 공정"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   3
               Left            =   2985
               TabIndex        =   11
               Top             =   90
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "패턴 공정"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin VB.TextBox txtCode 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFC0C0&
            Height          =   300
            Left            =   1305
            MaxLength       =   20
            TabIndex        =   2
            Top             =   60
            Width           =   1035
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Left            =   1305
            MaxLength       =   20
            TabIndex        =   6
            Top             =   720
            Width           =   3900
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   60
            TabIndex        =   1
            Top             =   60
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "코   드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   60
            TabIndex        =   5
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "패턴 설명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   4
            Left            =   60
            TabIndex        =   3
            Top             =   390
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "가   공"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Label lblName 
            Caption         =   "전체"
            ForeColor       =   &H000000C0&
            Height          =   1515
            Left            =   5310
            TabIndex        =   20
            Top             =   855
            Width           =   2910
         End
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   510
         Left            =   1380
         TabIndex        =   19
         Top             =   240
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
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10200
      TabIndex        =   17
      Top             =   6690
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Label lblCount 
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
      Height          =   225
      Left            =   75
      TabIndex        =   21
      Top             =   6915
      Width           =   3120
   End
End
Attribute VB_Name = "frmPatternCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_sFlag As String * 1

Dim m_bSortForward As Boolean

Private Sub cboWork_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Call NextFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click(Index As Integer)
    Dim sProcess$, sProcessID$
    Dim i%, nRow%
    
    If Index = 0 Then
        With grdProcess
            sProcess = .TextMatrix(.Row, 1)
            sProcessID = Format(.TextMatrix(.Row, 2), "0000")
        End With
        
        With grdPattern
            .AddItem "" & vbTab & sProcess & vbTab & sProcessID, .Row + 1
        
            nRow = .Row + 1
            
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, 0) = i
            Next i
            
            .Row = nRow
        End With
        
    Else
        With grdPattern
            If .Rows = .FixedRows Then Exit Sub
        
            If .Row < .FixedRows Then
                'MsgBox "리스트에서 목록을 선택하십시오", vbInformation
                Exit Sub
            End If
            
            nRow = .Row
            
            .RemoveItem (.Row)
            
            If .Rows > .FixedRows Then
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, 0) = i
                Next i
                
                .Row = nRow - 1
                If .Row = 0 Then
                    .Row = 1
                End If
            End If
            
        End With
    End If
End Sub

Private Sub ClearData()
    Dim i%
    
    txtCode = ""
    txtName = ""
    
    With grdPattern
        For i = .Rows - 1 To 1 Step -1
            .RemoveItem (i)
        Next i
    End With
    
    grdData(1).Rows = grdData(1).FixedRows
    
    
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
            pnlMsg.Caption = LoadResString(121)
            'txtCode.Locked = False
            txtName.SetFocus
    '---------------------------------------------------------------------------
        Case ID_UPDATE '[2] 수정
            If grdData(0).Rows = grdData(0).FixedRows Then Exit Sub
            m_sFlag = ID_UPDATE
            Call ChangeMode(Me, False)
            pnlMsg.Caption = LoadResString(122)
            
            txtCode.Locked = True
            txtName.SetFocus
    '---------------------------------------------------------------------------
        Case ID_DELETE '[3] 삭제
            If grdData(0).Rows = grdData(0).FixedRows Then Exit Sub
    
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                m_sFlag = ID_DELETE
            End If
            
            If SaveData Then
                Call FillGrid
            End If
    '---------------------------------------------------------------------------
        Case ID_SAVE  '[4] 저장
            If SaveData Then
                Call FillGrid
                Call ChangeMode(Me, True)
                m_sFlag = ""
                txtCode.Locked = False
            End If
        Case ID_CANCEL '[5] 취소
            m_sFlag = ""
            Call ChangeMode(Me, True)
            If grdData(0).Rows > 1 Then
                Call ShowData
            Else
                Call ClearData
            End If
            
    End Select
    
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmPatternCode.CmdOperate_Click", Err.Description)
'    MsgBox "[" & Err.Number & "]" & ":" & Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11985, 7845
    
    Call InitGrid
    Call SetOperate(Me)
    
    Call FillCombo
    Call FillGrid
    Call FillGridProcess
    txtCode.Locked = True
    
    lblName.Caption = "● 왼쪽 리스트 목록에서 공정을 " & vbCrLf & _
                                 "    선택한 다음 화살표 버튼으로 " & vbCrLf & _
                                 "    오른쪽으로 옮기십시오"
End Sub

Private Sub InitGrid()
    
    Call SetVSFlexGrid(grdData(0))
    Call SetVSFlexGrid(grdData(1))
    Call SetVSFlexGrid(grdProcess)
    Call SetVSFlexGrid(grdPattern)
    
    With grdData(0)
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 4
        
        .TextArray(1) = "코드":         .ColWidth(1) = 450:         .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "패턴설명":     .ColWidth(2) = 2430:        .ColAlignment(2) = flexAlignLeftCenter
        .Redraw = flexRDDirect
    End With
    
    With grdData(1)
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 2
        
        .TextArray(1) = "공정  순위":   .ColWidth(1) = 8050:        .ColAlignment(1) = flexAlignLeftCenter
        .Redraw = flexRDDirect
    End With
    
    
    With grdProcess
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 3
        
        .TextArray(1) = "공정명":   .ColWidth(1) = 1500:        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "공정ID":   .ColWidth(2) = 0:           .ColAlignment(2) = flexAlignLeftCenter
        .Redraw = flexRDDirect
        '.HighLight = flexHighlightAlways
    End With
    
    
    With grdPattern
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 3
        
        .TextArray(1) = "공정명":   .ColWidth(1) = 1500:        .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "공정ID":   .ColWidth(2) = 0:           .ColAlignment(2) = flexAlignLeftCenter
        .Redraw = flexRDDirect
        '.HighLight = flexHighlightAlways
    End With
    
    
End Sub
' 가공관리
Private Sub FillCombo()
    Dim oCode As PlusLib2.CCode
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon
    oCode.CodeType = CD_WORK
    Set rs = oCode.Getcode()

    Set oCode = Nothing

    With cboWork
        .Clear
        Do Until rs.EOF
            .AddItem CheckNull(rs!workname)
            .ItemData(.NewIndex) = rs!workid
            rs.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub

ErrHandler:
    
    Call ErrorBox(Err.Number, "frmPatternCode.FillCombo", Err.Description)
    Err.Clear
    Set oCode = Nothing

End Sub

Private Sub FillGrid()
    Dim oPattern As PlusLib2.CPattern
    Dim rs As ADODB.Recordset
    Dim lNowRow&
    
    On Error GoTo ErrHandler
    
    Set oPattern = New PlusLib2.CPattern
    oPattern.Connection = g_adoCon
    
    Set rs = oPattern.GetPattern
    Set oPattern = Nothing
    
    With grdData(0)
        .Redraw = False
        If .Rows > .FixedRows Then
            lNowRow = .Row
            .Rows = 1
        Else
            lNowRow = 1
        End If
        
            Do Until rs.EOF
                .AddItem CStr(.Rows) & vbTab & rs!PatternID & vbTab & CheckNull(rs!Pattern) & vbTab & rs!workid
                rs.MoveNext
            Loop
            
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            If .Rows > lNowRow Then
                .Row = lNowRow
            Else
                .Row = .Rows - 1
            End If
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
            Call ShowData
        Else
            .HighLight = flexHighlightNever
            Call ClearData
        End If
        
        .Redraw = True
    End With
    lblCount.Caption = LoadResString(250) & grdData(0).Rows - 1 & " 건"
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrHandler:
    'MsgBox "[" & Err.Number & "]" & ":" & Err.Description, vbCritical
    Call ErrorBox(Err.Number, "frmPatternCode.FillGrid", Err.Description)
    Err.Clear
    Set rs = Nothing
    Set oPattern = Nothing
End Sub


Private Sub grdData_RowColChange(Index As Integer)
    If Index = 0 Then
        Call ShowData
    End If
End Sub


Private Sub ShowData()
    Dim oPattern As PlusLib2.CPattern
    Dim rs As ADODB.Recordset
    Dim iLoop%, i%
    Dim sProcess$
    
    On Error GoTo ErrHandler
    
    Set oPattern = New PlusLib2.CPattern
    oPattern.Connection = g_adoCon

    Set rs = oPattern.GetPatternSub(grdData(0).TextMatrix(grdData(0).Row, 1))
    Set oPattern = Nothing

    With grdPattern
        .Redraw = flexRDNone
        For i = .Rows - 1 To 1 Step -1
            .RemoveItem (i)
        Next i
        .Redraw = flexRDDirect
    End With
    
    With grdData(1)
        .Redraw = flexRDNone
        .Rows = 1
        
        Do Until rs.EOF
            sProcess = sProcess & "→" & "[" & CheckNull(rs!Process) & "]"
            
            With grdPattern
                .AddItem CStr(grdPattern.Rows) & vbTab & CheckNull(rs!Process) & vbTab & CheckNull(rs!processid)
            End With
            
            rs.MoveNext
        Loop
        
        .AddItem "" & vbTab & Mid(sProcess, 2)
        
        .Redraw = flexRDDirect
    End With
    
    If grdPattern.Rows > grdPattern.FixedRows Then
        grdPattern.Row = grdPattern.FixedRows
    End If
    
    With grdData(0)
        txtCode = .TextMatrix(.Row, 1)
        txtName = .TextMatrix(.Row, 2)
        cboWork.ListIndex = FindComboBox(cboWork, CLng(.TextMatrix(.Row, 3)))
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrHandler:
    'MsgBox "[" & Err.Number & "]" & ":" & Err.Description, vbCritical
    Call ErrorBox(Err.Number, "frmPatternCode.ShowData", Err.Description)
    Set rs = Nothing
    Set oPattern = Nothing

End Sub

Private Sub FillGridProcess()
    Dim oProcess As PlusLib2.cprocess
    Dim rs As ADODB.Recordset
    Dim iLoop
    
    On Error GoTo ErrHandler
    
    Set oProcess = New PlusLib2.cprocess
    oProcess.Connection = g_adoCon
    Set rs = oProcess.GetProcess()
    Set oProcess = Nothing
    
    With grdProcess
        .Redraw = flexRDNone
        For iLoop = 1 To rs.RecordCount
            .AddItem CStr(iLoop) & vbTab & rs!Process & vbTab & CLng(rs!processid)
            rs.MoveNext
        Next iLoop
        
        .Redraw = flexRDDirect
    End With
    
    rs.Close
    Set rs = Nothing
    Exit Sub

ErrHandler:
    'MsgBox CStr(Err.Number) & ":" & Err.Description, vbCritical
    Call ErrorBox(Err.Number, "frmPatternCode.FillGridProcess", Err.Description)
    Err.Clear
    rs.Close
    Set rs = Nothing
    Set oProcess = Nothing
    
End Sub

Private Function CheckData() As Long
    Dim iLoop As Integer

    CheckData = -1
    If grdPattern.Rows = 1 Then
        MsgBox "추가된 공정이 없습니다", vbInformation
        CheckData = -1
        Exit Function
    End If

    CheckData = grdPattern.Rows - 1

End Function

Private Sub SetNewData(SetPatternData As PlusLib2.tPattern, SetSubData() As PlusLib2.TSubPattern, nSeq As Integer)
    Dim iLoop%

    With SetPatternData
        .sPatternID = txtCode
        .sPattern = txtName
        .sWorkID = Format(cboWork.ItemData(cboWork.ListIndex), "0000")
    End With

    For iLoop = 0 To nSeq
        With SetSubData(iLoop)
            .sPatternID = txtCode
            .sProcessID = Format(grdPattern.TextMatrix(iLoop + 1, 2), "0000")
            '.sProcessID = Format(grdPattern.ItemData(iLoop), "00")
            .nPatternSeq = iLoop + 1
        End With
    Next iLoop
End Sub

Private Function SaveData() As Boolean
    Dim oPattern As PlusLib2.CPattern
    Dim NewPattern   As PlusLib2.tPattern
    Dim SubPattern() As PlusLib2.TSubPattern
    Dim nSeq  As Integer
    Dim i%
    
    On Error GoTo ErrHandler
    
    SaveData = False
    ' 코드 Check
    If m_sFlag = ID_ADDNEW Then
        With grdData(0)
            For i = 1 To .Rows - 1
                If txtCode = .TextMatrix(i, 1) Then
                    MsgBox LoadResString(205), vbInformation
                    txtCode.SetFocus
                    Exit Function
                End If
            Next i
        End With
    End If
    
    If CheckData = -1 Then Exit Function
    
    nSeq = CheckData
    
    If nSeq > 0 Then ReDim SubPattern(nSeq - 1)
    
    If m_sFlag <> ID_DELETE Then
    
        Call SetNewData(NewPattern, SubPattern, nSeq - 1)
    
        If SubPattern(0).sProcessID <> "0401" Then
            Call MessageBox("모든 공정 패턴은 '배색'공정으로 시작되어야 합니다." & vbCrLf & vbCrLf & "다시 확인해 주십시오")
            Exit Function
        End If
    End If
    
    Set oPattern = New PlusLib2.CPattern
    oPattern.Connection = g_adoCon
    oPattern.UserName = g_sUserName
    
    Select Case m_sFlag
        Case ID_ADDNEW
            SaveData = oPattern.AddNewPattern(NewPattern, SubPattern, nSeq - 1)
        Case ID_UPDATE
            SaveData = oPattern.UpdatePattern(NewPattern, SubPattern, nSeq - 1)
        Case ID_DELETE
            SaveData = oPattern.DeletePattern(grdData(0).TextMatrix(grdData(0).Row, 1))
    End Select
    
    Set oPattern = Nothing
    
    Exit Function

ErrHandler:

    Call ErrorBox(Err.Number, "frmPatternCode.Save", Err.Description)
    Err.Clear
    Set oPattern = Nothing

End Function

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyPress(txtCode, KeyAscii, True, 2)
End Sub

Private Sub txtName_GotFocus()
    Call GotFocusText(txtName)
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call NextFocus
    End If
End Sub
