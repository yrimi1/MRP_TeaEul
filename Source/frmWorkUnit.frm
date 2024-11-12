VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWorkUnit 
   ClientHeight    =   9255
   ClientLeft      =   405
   ClientTop       =   1095
   ClientWidth     =   11850
   Icon            =   "frmWorkUnit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin VB.TextBox txtBatJaNo 
      Height          =   315
      Left            =   8760
      TabIndex        =   35
      Top             =   6960
      Width           =   975
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   345
      Index           =   5
      Left            =   8760
      TabIndex        =   34
      Top             =   6540
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   609
      _Version        =   196609
      Caption         =   "밧자기 번호"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdMove 
      Height          =   795
      Index           =   0
      Left            =   10890
      TabIndex        =   30
      Top             =   7470
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1402
      _Version        =   196609
      Caption         =   "위"
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdMove 
      Height          =   795
      Index           =   1
      Left            =   9810
      TabIndex        =   31
      Top             =   7470
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1402
      _Version        =   196609
      Caption         =   "아래"
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   690
      Left            =   8400
      TabIndex        =   32
      Tag             =   "PERM_ADDNEW"
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "취소"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   690
      Left            =   6660
      TabIndex        =   33
      Tag             =   "PERM_ADDNEW"
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "저장"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdWorkUnit 
      Height          =   1905
      Left            =   0
      TabIndex        =   29
      Top             =   6450
      Width           =   8655
      _cx             =   15266
      _cy             =   3360
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
   Begin Threed.SSCommand cmdAdd 
      Height          =   795
      Left            =   9810
      TabIndex        =   25
      Tag             =   "PERM_ADDNEW"
      Top             =   6510
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1402
      _Version        =   196609
      Caption         =   "추가"
      Alignment       =   8
      PictureAlignment=   6
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   420
      TabIndex        =   21
      Top             =   3210
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
         TabIndex        =   22
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
         TabIndex        =   23
         Top             =   120
         Width           =   270
      End
   End
   Begin Threed.SSFrame frmSearch 
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1614
      _Version        =   196609
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   6600
         TabIndex        =   26
         Top             =   495
         Width           =   1905
      End
      Begin VB.ComboBox cboProcess 
         Height          =   300
         Left            =   8610
         Style           =   2  '드롭다운 목록
         TabIndex        =   18
         Top             =   495
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   10950
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   4
         ToolTipText     =   "자료 저장"
         Top             =   60
         Width           =   780
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   6600
         TabIndex        =   3
         Top             =   75
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   2820
         TabIndex        =   2
         Top             =   495
         Width           =   1905
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   2850
         TabIndex        =   1
         Top             =   75
         Width           =   1905
      End
      Begin Threed.SSPanel pnlOrder 
         Height          =   795
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1402
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   120
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   480
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   9
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   4785
         TabIndex        =   10
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
         Left            =   1440
         TabIndex        =   11
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
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
            TabIndex        =   12
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   4770
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   495
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
         Left            =   5220
         TabIndex        =   14
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
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
            TabIndex        =   15
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   5220
         TabIndex        =   16
         Top             =   495
         Width           =   1320
         _ExtentX        =   2328
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
            Caption         =   "카드번호"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   8610
         TabIndex        =   27
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
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
            Caption         =   "대기공정"
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   1185
         End
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   5505
      Left            =   0
      TabIndex        =   19
      Top             =   930
      Width           =   11835
      _cx             =   20876
      _cy             =   9710
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
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10140
      TabIndex        =   20
      Top             =   8460
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdDel 
      Height          =   795
      Left            =   10890
      TabIndex        =   24
      Tag             =   "PERM_ADDNEW"
      Top             =   6510
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1402
      _Version        =   196609
      Caption         =   "삭제"
      Alignment       =   8
      PictureAlignment=   6
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmWorkUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bloading As Boolean

Private Sub chkSearch_Click(Index As Integer)
    If Index >= 1 And Index <= 4 Then
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
    Else
        If chkSearch(Index).Value = vbChecked Then
            cboProcess.Enabled = True
            cboProcess.SetFocus
        Else
            cboProcess.Enabled = False
        End If
    End If
End Sub

Private Function ExistCard(nRow As Integer, sCardID As String, sSplitID As String) As Boolean
    Dim i%
    
    ExistCard = False
    
    With grdWorkUnit
        For i = 1 To .Rows - .FixedRows
            If sCardID = .TextMatrix(i, 3) And sSplitID = .TextMatrix(i, 4) Then
                ExistCard = True
                Exit Function
            End If
        Next i
    End With
End Function

Private Sub cmdAdd_Click()
    Dim i%, iRow%, sWorkUnit$
    
    If grdData.Rows = grdData.FixedRows Then Exit Sub
    
    With grdData
        If grdData.TextMatrix(grdData.Row, 12) = "보류" Then
            MsgBox "보류중인 카드는 작업순서를 변경할 수 없습니다.!!", vbInformation + vbOKOnly
            Exit Sub
        End If
    End With
    
    With grdWorkUnit
        If .Rows > .FixedRows And .TextMatrix(.Row, 8) <> grdData.TextMatrix(grdData.Row, 11) Then
            MsgBox "대기공정이 틀린 카드는 작업순서를 변경할 수 없습니다.!!", vbInformation + vbOKOnly
            Exit Sub
        End If
    End With
        
    With grdData
        If .IsSubtotal(grdData.Row) = True Then
            sWorkUnit = .TextMatrix(.Row + 1, 14)
            
            For i = .FixedRows To .Rows - 1
                If sWorkUnit = .TextMatrix(i, 14) Then
                    If ExistCard(i, .TextMatrix(i, 6), .TextMatrix(i, 7)) Then Exit For
                    If grdData.TextMatrix(i, 12) = "보류" Then Exit For
                    
                    grdWorkUnit.Rows = grdWorkUnit.Rows + 1
                    iRow = grdWorkUnit.Rows - 1
                
                    grdWorkUnit.TextMatrix(iRow, 0) = iRow
                    grdWorkUnit.TextMatrix(iRow, 1) = MakeCardID(.TextMatrix(i, 14), OM_EXPAND)
                    grdWorkUnit.TextMatrix(iRow, 2) = .TextMatrix(i, 1)
                    grdWorkUnit.TextMatrix(iRow, 3) = .TextMatrix(i, 6)
                    grdWorkUnit.TextMatrix(iRow, 4) = .TextMatrix(i, 7)
                    grdWorkUnit.TextMatrix(iRow, 5) = .TextMatrix(i, 8)
                    grdWorkUnit.TextMatrix(iRow, 6) = .TextMatrix(i, 9)
                    grdWorkUnit.TextMatrix(iRow, 7) = .TextMatrix(i, 13)
                    grdWorkUnit.TextMatrix(iRow, 8) = .TextMatrix(i, 11)
                    
                    txtBatJaNo = .TextMatrix(i, 13)
                End If
            Next i
        Else
            If ExistCard(.Row, .TextMatrix(.Row, 6), .TextMatrix(.Row, 7)) Then Exit Sub
            
            grdWorkUnit.Rows = grdWorkUnit.Rows + 1
            iRow = grdWorkUnit.Rows - 1
        
            grdWorkUnit.TextMatrix(iRow, 0) = iRow
            grdWorkUnit.TextMatrix(iRow, 1) = MakeCardID(.TextMatrix(.Row, 14), OM_EXPAND)
            grdWorkUnit.TextMatrix(iRow, 2) = .TextMatrix(.Row, 1)
            grdWorkUnit.TextMatrix(iRow, 3) = .TextMatrix(.Row, 6)
            grdWorkUnit.TextMatrix(iRow, 4) = .TextMatrix(.Row, 7)
            grdWorkUnit.TextMatrix(iRow, 5) = .TextMatrix(.Row, 8)
            grdWorkUnit.TextMatrix(iRow, 6) = .TextMatrix(.Row, 9)
            grdWorkUnit.TextMatrix(iRow, 7) = .TextMatrix(.Row, 13)
            grdWorkUnit.TextMatrix(iRow, 8) = .TextMatrix(.Row, 11)
            
            txtBatJaNo = .TextMatrix(.Row, 13)
        End If
        If grdWorkUnit.Rows > grdWorkUnit.FixedRows Then
            grdWorkUnit.HighLight = flexHighlightAlways
            grdWorkUnit.Row = grdWorkUnit.FixedRows
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    grdWorkUnit.Rows = grdWorkUnit.FixedRows
End Sub

Private Sub cmdDel_Click()
    Dim i%
    
    With grdWorkUnit
        If .Rows = .FixedRows Then Exit Sub
        .RemoveItem .Row
        
        For i = 1 To .Rows - .FixedRows
            .TextMatrix(i, 0) = i
        Next i
        
        If .Rows > .FixedRows Then
            txtBatJaNo = .TextMatrix(.Rows - 1, 7)
        Else
            txtBatJaNo = ""
        End If
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

Private Sub cmdMove_Click(Index As Integer)
    Dim i%, iRow%
    Dim sWorkUnit$, WorkUnit() As String
    
    If grdWorkUnit.Rows = grdWorkUnit.FixedRows Then Exit Sub
    ReDim WorkUnit(7)
    If Index = 0 Then
        With grdWorkUnit
            If .Row = .FixedRows Then Exit Sub
            iRow = .Row
            
            WorkUnit(1) = .TextMatrix(.Row - 1, 1)
            WorkUnit(2) = .TextMatrix(.Row - 1, 2)
            WorkUnit(3) = .TextMatrix(.Row - 1, 3)
            WorkUnit(4) = .TextMatrix(.Row - 1, 4)
            WorkUnit(5) = .TextMatrix(.Row - 1, 5)
            WorkUnit(6) = .TextMatrix(.Row - 1, 6)
            WorkUnit(7) = .TextMatrix(.Row - 1, 7)
            
            .TextMatrix(.Row - 1, 1) = .TextMatrix(.Row, 1)
            .TextMatrix(.Row - 1, 2) = .TextMatrix(.Row, 2)
            .TextMatrix(.Row - 1, 3) = .TextMatrix(.Row, 3)
            .TextMatrix(.Row - 1, 4) = .TextMatrix(.Row, 4)
            .TextMatrix(.Row - 1, 5) = .TextMatrix(.Row, 5)
            .TextMatrix(.Row - 1, 6) = .TextMatrix(.Row, 6)
            .TextMatrix(.Row - 1, 7) = .TextMatrix(.Row, 7)
            
            .TextMatrix(.Row, 1) = WorkUnit(1)
            .TextMatrix(.Row, 2) = WorkUnit(2)
            .TextMatrix(.Row, 3) = WorkUnit(3)
            .TextMatrix(.Row, 4) = WorkUnit(4)
            .TextMatrix(.Row, 5) = WorkUnit(5)
            .TextMatrix(.Row, 6) = WorkUnit(6)
            .TextMatrix(.Row, 7) = WorkUnit(7)
                        
            .Row = iRow - 1
        End With
    Else
        With grdWorkUnit
            If .Row = .Rows - 1 Then Exit Sub
            iRow = .Row
            
            WorkUnit(1) = .TextMatrix(.Row + 1, 1)
            WorkUnit(2) = .TextMatrix(.Row + 1, 2)
            WorkUnit(3) = .TextMatrix(.Row + 1, 3)
            WorkUnit(4) = .TextMatrix(.Row + 1, 4)
            WorkUnit(5) = .TextMatrix(.Row + 1, 5)
            WorkUnit(6) = .TextMatrix(.Row + 1, 6)
            WorkUnit(7) = .TextMatrix(.Row + 1, 7)
            
            .TextMatrix(.Row + 1, 1) = .TextMatrix(.Row, 1)
            .TextMatrix(.Row + 1, 2) = .TextMatrix(.Row, 2)
            .TextMatrix(.Row + 1, 3) = .TextMatrix(.Row, 3)
            .TextMatrix(.Row + 1, 4) = .TextMatrix(.Row, 4)
            .TextMatrix(.Row + 1, 5) = .TextMatrix(.Row, 5)
            .TextMatrix(.Row + 1, 6) = .TextMatrix(.Row, 6)
            .TextMatrix(.Row + 1, 7) = .TextMatrix(.Row, 7)
            
            .TextMatrix(.Row, 1) = WorkUnit(1)
            .TextMatrix(.Row, 2) = WorkUnit(2)
            .TextMatrix(.Row, 3) = WorkUnit(3)
            .TextMatrix(.Row, 4) = WorkUnit(4)
            .TextMatrix(.Row, 5) = WorkUnit(5)
            .TextMatrix(.Row, 6) = WorkUnit(6)
            .TextMatrix(.Row, 7) = WorkUnit(7)
            
            .Row = iRow + 1
        End With
    End If
End Sub

Private Sub cmdSave_Click()
    Dim oWorkUnit As PlusLib2.CWorkUnit
    Dim tWork() As PlusLib2.TWorkUnit

    Dim i%
        
    On Error GoTo ErrHandler
    
    If grdWorkUnit.Rows = grdWorkUnit.FixedRows Then Exit Sub
    
    If Not CheckBatJaNo Then Exit Sub
    
    With grdWorkUnit
        ReDim tWork(.Rows - .FixedRows - 1)
        For i = .FixedRows To .Rows - 1
            tWork(i - 1).sCardID = MakeCardID(.TextMatrix(i, 3), OM_REDUCE)
            tWork(i - 1).sSplitID = .TextMatrix(i, 4)
            If Len(txtBatJaNo) > 0 Then
                tWork(i - 1).sBatJaNo = txtBatJaNo
            End If
        Next i
    End With
    
    Set oWorkUnit = New PlusLib2.CWorkUnit
    oWorkUnit.Connection = g_adoCon
    oWorkUnit.UserName = g_sUserName
    
    If oWorkUnit.ModifyWorkUnit(tWork()) Then
        grdWorkUnit.Rows = grdWorkUnit.FixedRows
        Call FillGridData
    End If
    Set oWorkUnit = Nothing
    Exit Sub
ErrHandler:
    Set oWorkUnit = Nothing
    Call ErrorBox(Err.Number, "frmWorkUnit.CmdSave_Click", Err.Description)
End Sub

Private Function CheckBatJaNo() As Boolean
    Dim oWorkUnit As PlusLib2.CWorkUnit
    
    Set oWorkUnit = New PlusLib2.CWorkUnit
    oWorkUnit.Connection = g_adoCon
    
    CheckBatJaNo = True
    If Len(Trim(txtBatJaNo)) = 0 Then Exit Function
    If Not oWorkUnit.GetBatJaNo(Trim(txtBatJaNo)) Then
        MsgBox "해당 밧자번호가 등록되어 있지 않습니다. 밧자번호를 등록하여 주십시오", vbCritical
        txtBatJaNo = ""
        CheckBatJaNo = False
    End If
End Function

Private Sub cmdSearch_Click()
    Call FillGridData
End Sub



Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11970, 9660
    
    Call SetOperate(Me)
    Call InitGrid
    Call MakeProcessCombo
    
    For i = 1 To 2
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        txtSearch(i).Enabled = False
    Next i
    txtSearch(3).Enabled = False
    txtSearch(4).Enabled = False
    cboProcess.Enabled = False
    
    pnlProgress.Visible = False
    
    cmdAdd.Picture = LoadResPicture("ADDNEW", vbResIcon)
    cmdDel.Picture = LoadResPicture("DELETE", vbResIcon)
    cmdMove(0).Picture = LoadResPicture("UP", vbResIcon)
    cmdMove(1).Picture = LoadResPicture("DOWN", vbResIcon)
    
End Sub

Private Sub grdData_DblClick()
    With grdData
        If .Row < .FixedRows Then Exit Sub

        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If
    End With
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
    ElseIf KeyAscii = vbKeyReturn And Index >= 3 Then
        Call NextFocus
    End If
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdData
        .Cols = 15

        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        
        .RowHeight(0) = 450
        
        .TextArray(0) = " ":                          .ColWidth(0) = 250
        .TextArray(1) = "작업단위" & vbCrLf & "순번": .ColWidth(1) = 1300:            .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "거래처":       .ColWidth(2) = 1000:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "품명":         .ColWidth(3) = 2000:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "관리번호":     .ColWidth(4) = 1350:            .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "OrderNo":      .ColWidth(5) = 0:               .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "카드번호":     .ColWidth(6) = 1000:            .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "분할" & vbCrLf & "번호":     .ColWidth(7) = 600:            .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "절수":         .ColWidth(8) = 500:            .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "카드" & vbCrLf & "수량":     .ColWidth(9) = 600:            .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "완료공정":    .ColWidth(10) = 800:           .ColAlignment(10) = flexAlignCenterCenter
        .TextArray(11) = "대기공정":    .ColWidth(11) = 800:           .ColAlignment(11) = flexAlignCenterCenter
        .TextArray(12) = "카드상태":    .ColWidth(12) = 800:           .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "BatJaNo":     .ColWidth(13) = 800:           .ColAlignment(13) = flexAlignCenterCenter
        .TextArray(14) = "작업단위":    .ColWidth(14) = 0
        
        .ColFormat(8) = "#,##0"
        .ColFormat(9) = "#,##0"
        
        For i = .FixedCols To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 0
        .ExtendLastCol = True
        .Editable = flexEDKbd
        .Redraw = flexRDDirect
    End With
    
    With grdWorkUnit
        .Cols = 9
        
        Call SetVSFlexGrid(grdWorkUnit)
        
        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        
        .TextArray(0) = " "
        .TextArray(1) = "작업단위":         .ColWidth(1) = 1550:             .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "현재순위":         .ColWidth(2) = 1500:             .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "카드번호":         .ColWidth(3) = 1350:             .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "분할번호":         .ColWidth(4) = 1000:             .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "절수":             .ColWidth(5) = 1000:             .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "수량":             .ColWidth(6) = 1000:             .ColAlignment(6) = flexAlignRightCenter
        .TextArray(7) = "BatJaNo":            .ColWidth(7) = 1000:             .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "대기공정":            .ColWidth(8) = 0:             .ColAlignment(7) = flexAlignCenterCenter
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub MakeProcessCombo()
    Dim oWorkUnit As PlusLib2.CWorkUnit
    Dim rs As Recordset

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True
    
    Set oWorkUnit = New PlusLib2.CWorkUnit
    oWorkUnit.Connection = g_adoCon

    Set rs = oWorkUnit.GetProcess(1)
    Set oWorkUnit = Nothing

    With cboProcess
        .Clear

        Do Until rs.EOF
            .AddItem CStr(rs!Process)
            .ItemData(.NewIndex) = CLng(Left(rs!ProcessID, 2))
            
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing

        If .ListCount > 0 Then .ListIndex = 0
    End With

    m_bloading = False
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oWorkUnit = Nothing
    Screen.MousePointer = vbDefault
    m_bloading = False
    Call ErrorBox(Err.Number, "frmWorkUnit.MakeProcessCombo", Err.Description)
End Sub

Private Sub FillGridData()
    Dim oWorkUnit As PlusLib2.CWorkUnit
    Dim rs As ADODB.Recordset
    Dim i%, nTop%
    Dim sWorkUnit$, sOrderID$
    
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    Set oWorkUnit = New PlusLib2.CWorkUnit
    oWorkUnit.Connection = g_adoCon
    
    Set rs = oWorkUnit.GetOrder(IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, IIf(chkSearch(2) = vbChecked, 1, 0), txtSearch(2).Tag, _
                                 IIf(chkSearch(3) = vbChecked, IIf(optOrder(1).Value = True, 1, 2), 0), txtSearch(3), _
                                 IIf(chkSearch(4) = vbChecked, 1, 0), txtSearch(4), _
                                 IIf(chkSearch(5) = vbChecked, 1, 0), Format(Left(cboProcess.ItemData(cboProcess.ListIndex), 2), "00"))
    Set oWorkUnit = Nothing
        
    With grdData
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
                            
            If rs!WorkUnitId <> sWorkUnit Then
                .AddItem "" & vbTab & MakeWorkUnitID(rs!WorkUnitId, OM_EXPAND) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                    "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                    "" & vbTab & "" & vbTab & ""
                
                Call DoFlexGridGroup(grdData, .Rows - 1, 1)
            End If
        
            .AddItem "" & vbTab & rs!WorkUnitSeq & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                    MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & _
                    rs!SplitID & vbTab & rs!Roll & vbTab & rs!Qty & vbTab & rs!CompProc & vbTab & rs!WaitProc & vbTab & _
                    rs!UseClss & vbTab & rs!BatJaNO & vbTab & rs!WorkUnitId
            
            If rs!UseClss = "보류" Then
                .Cell(flexcpBackColor, .Rows - 1, 6, .Rows - 1, 7) = vbRed
                .Cell(flexcpForeColor, .Rows - 1, 6, .Rows - 1, 7) = vbWhite
            ElseIf rs!UseClss = "작업" Then
                .Cell(flexcpBackColor, .Rows - 1, 6, .Rows - 1, 7) = vbBlue
                .Cell(flexcpForeColor, .Rows - 1, 6, .Rows - 1, 7) = vbWhite
            End If
            
            lblCount = CStr(i) & " / " & CStr(rs.RecordCount) & "  (" & Format((i / rs.RecordCount) * 100, "00.0") & " %)"
            proProgress.Value = CInt((i / rs.RecordCount) * 100)
            
            sOrderID = rs!OrderID
            sWorkUnit = rs!WorkUnitId
            
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
            .Rows = .FixedRows
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    pnlProgress.Visible = False
    Exit Sub

ErrHandler:
    Set oWorkUnit = Nothing
    Set rs = Nothing
    pnlProgress.Visible = False
    Call ErrorBox(Err.Number, "frmWorkUnit.FillGridData", Err.Description)
End Sub


Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, iRow As Integer, iLvl As Integer)
    With oFlex
        ' Set the row as a group
        .IsSubtotal(iRow) = True
        ' Set the indentation level of the group
        .RowOutlineLevel(iRow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
        Case 1, 2
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = COLOR_GRIDROW
'            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HFFFFC0    '&HE0E0E0
        End Select
    End With
End Sub

