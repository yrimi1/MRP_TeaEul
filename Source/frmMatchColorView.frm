VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMatchColorView 
   ClientHeight    =   9270
   ClientLeft      =   1650
   ClientTop       =   2835
   ClientWidth     =   15180
   Icon            =   "frmMatchColorView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15180
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전일"
      Height          =   315
      Index           =   0
      Left            =   60
      MousePointer    =   99  '사용자 정의
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   390
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   690
      MousePointer    =   99  '사용자 정의
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   390
      Width           =   615
   End
   Begin VB.TextBox txtCardID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10770
      MaxLength       =   12
      TabIndex        =   22
      Top             =   60
      Width           =   1605
   End
   Begin VB.ComboBox cboTeamID 
      Height          =   300
      Left            =   10785
      Style           =   2  '드롭다운 목록
      TabIndex        =   10
      Top             =   405
      Width           =   1590
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Left            =   4590
      TabIndex        =   9
      Top             =   450
      Width           =   1275
   End
   Begin VB.TextBox txtCustomID 
      Height          =   285
      Left            =   7500
      TabIndex        =   8
      Top             =   60
      Width           =   1605
   End
   Begin VB.TextBox txtArticleID 
      Height          =   300
      Left            =   7500
      TabIndex        =   7
      Top             =   405
      Width           =   1605
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   690
      Left            =   14070
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   4
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   1065
   End
   Begin VB.Frame fraOrder 
      Height          =   510
      Left            =   3330
      TabIndex        =   1
      Top             =   -60
      Width           =   2865
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   1530
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   1155
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13500
      TabIndex        =   0
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdSum 
      Height          =   900
      Left            =   0
      TabIndex        =   5
      Top             =   6030
      Width           =   15180
      _cx             =   26776
      _cy             =   1587
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
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
      Height          =   5265
      Left            =   0
      TabIndex        =   6
      Top             =   750
      Width           =   15180
      _cx             =   26776
      _cy             =   9287
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
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
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   9510
      TabIndex        =   11
      Top             =   405
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   196610
      Caption         =   "작 업 조"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkTeamID 
         Caption         =   "작 업 조"
         Height          =   180
         Left            =   75
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   60
         Width           =   960
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   5
      Left            =   3330
      TabIndex        =   13
      Top             =   450
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   196610
      Caption         =   "기    계"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkOrder 
         Caption         =   "Order No."
         Height          =   180
         Left            =   75
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   60
         Width           =   1125
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   285
      Index           =   6
      Left            =   6240
      TabIndex        =   15
      Top             =   60
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   196610
      Caption         =   "작 업 조"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkCustom 
         Caption         =   "거 래 처"
         Height          =   180
         Left            =   75
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   60
         Width           =   960
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   3
      Left            =   5910
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   450
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196610
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
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   4
      Left            =   9150
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   60
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      _Version        =   196610
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
      Index           =   10
      Left            =   6240
      TabIndex        =   19
      Top             =   405
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   196610
      Caption         =   "기    계"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkArticleID 
         Caption         =   "품     명"
         Height          =   180
         Left            =   75
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   60
         Width           =   1050
      End
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   300
      Index           =   5
      Left            =   9150
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   405
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   196610
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
      Height          =   285
      Index           =   8
      Left            =   9510
      TabIndex        =   23
      Top             =   60
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   196610
      Caption         =   "기    계"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkCardSearch 
         Caption         =   "카드번호"
         Height          =   180
         Left            =   75
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   60
         Width           =   1050
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdDetail 
      Height          =   1575
      Left            =   0
      TabIndex        =   25
      Top             =   6930
      Width           =   15180
      _cx             =   26776
      _cy             =   2778
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Index           =   1
      Left            =   1350
      TabIndex        =   26
      Top             =   420
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   131989504
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Index           =   0
      Left            =   1350
      TabIndex        =   27
      Top             =   60
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   131989504
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   3
      Left            =   60
      TabIndex        =   28
      Top             =   60
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   196610
      Caption         =   "실적 일자"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11790
      TabIndex        =   31
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmMatchColorView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const LIMIT_ROW = 23
Private Const LIMIT_WIDTH0 = 1365
Private Const LIMIT_WIDTH1 = 1890
Private Const LIMIT_WIDTH2 = 1270
Private Const LIMIT_WIDTH3 = 1740
Private Const LIMIT_WIDTH4 = 1200
Private Const LIMIT_WIDTH5 = 1755
Private Const LIMIT_WIDTH6 = 1780
Private Const LIMIT_WIDTH7 = 1740

Private m_iSortType As Integer
Private m_bloading  As Boolean
Private m_bSkip As Boolean

Private Type TParaValue
    sDate           As String
    eDate           As String
    nCheckOrderID   As Integer
    sOrderID        As String
    nCheckOrderNo   As Integer
    sOrderNO        As String
    nCheckCutom     As Integer
    sCustomID       As String
    nCheckArticle   As Integer
    sArticleID      As String
    nCheckTeam      As Integer
    sTeamID         As String
    nCheckCardID    As Integer
    sCardID         As String
End Type



Private Sub chkArticleID_Click()

    If chkArticleID.Value = vbChecked Then
        txtArticleID.Locked = False
        cmdFind(5).Enabled = True
    Else
        txtArticleID.Locked = True
        txtArticleID.Text = ""
        cmdFind(5).Enabled = False
    End If
End Sub

Private Sub chkCardSearch_Click()

    If chkCardSearch.Value = vbChecked Then
        txtCardID.Enabled = True
        
    Else
        txtCardID.Enabled = False
        txtCardID.Text = ""
    End If
    
End Sub

Private Sub chkCustom_Click()
    If chkCustom.Value = vbChecked Then
        txtCustomID.Locked = False
        cmdFind(4).Enabled = True
    Else
        txtCustomID.Locked = True
        txtCustomID.Text = ""
        cmdFind(4).Enabled = False
    End If
End Sub

Private Sub chkOrder_Click()
    If chkOrder.Value = vbChecked Then
        txtOrder.Locked = False
        cmdFind(3).Enabled = True
    Else
        txtOrder.Locked = True
        txtOrder.Text = ""
        cmdFind(3).Enabled = False
    End If
End Sub

Private Sub chkTeamID_Click()
    If chkTeamID.Value = vbChecked Then
        cboTeamID.Enabled = True
        
    Else
        cboTeamID.Enabled = False
        cboTeamID.ListIndex = -1
    End If
End Sub



Private Sub cmdPrint_Click()
    With grdData
        .Redraw = flexRDBuffered
        
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(.Rows - 1) = False
        
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "배 색  일 지"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpText, 1, 1, 1, 6) = "▶ 실적일 : " & Format(dtpDate(0), "YYYY/MM/DD") & " ~ " & Format(dtpDate(1), "YYYY/MM/DD")
        .Cell(flexcpText, 1, 12, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        .Cell(flexcpBackColor, 0, 0, 1, .Cols - 1) = vbWhite
        
        .ColWidth(0) = 0
        .ColWidth(1) = 400
        .ColWidth(2) = 700
        .ColWidth(3) = 1150
        .ColWidth(4) = 1500
        .ColWidth(5) = 1300
        .ColWidth(6) = 1400
        .ColWidth(7) = 1000
        .ColWidth(8) = 1500
        .ColWidth(9) = 500
        .ColWidth(10) = 800
        .ColWidth(11) = 900
        .ColWidth(12) = 800
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 700
        .ColWidth(16) = 700
        .ColWidth(17) = 700
        .ColWidth(18) = 600
        
        Call SetPrintMode(grdData, 1, True)
        .PrintGrid "태을염직", True, 2, 100, 500
        Call SetPrintMode(grdData, 1, False)

'        .GridLinesFixed = flexGridInset
        .RowHidden(0) = True
        .RowHidden(1) = True
        
        .RowHidden(.Rows - 1) = True
        .RowHidden(.Rows - 2) = True
        .RowHidden(.Rows - 3) = True

        .ColWidth(3) = 1400
        .ColWidth(4) = 2200
        .ColWidth(7) = 1500
        .ColWidth(11) = 1400
        .ColWidth(12) = 700
        .ColWidth(14) = 700
        .ColWidth(16) = 800
        .ColWidth(17) = 800
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpDate(0) = DateSerial(Year(Date), Month(Date), Day(Date) - 1)
            dtpDate(1) = DateSerial(Year(Date), Month(Date), Day(Date) - 1)
        Case 1
            dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
            dtpDate(1) = Date
    End Select
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub FillGridData()
    Dim adoCmd As ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim sTeamID As String
    Dim dSql_str As String
    Dim TParaValue As TParaValue
    Dim dTotRoll As Long, dTotQty As Long
    Dim dReWorkRoll As Long, dReWorkQty As Long
    
    
    If chkTeamID.Value = vbChecked Then
        dSql_str = "SELECT TeamID  FROM [mt_team] WHERE Team = '" & Trim(cboTeamID.Text) & "' "
        rs.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
        If rs.RecordCount = 1 Then
            sTeamID = Trim(rs(0))
        End If
        rs.Close
        Set rs = Nothing
    Else
        sTeamID = ""
    End If
    
    '------ Parameter 넘겨줄 값 Move

    With TParaValue
        If chkOrder.Value = vbChecked Then
            If optOrder(0).Value = True Then  'Order NO
                .nCheckOrderID = 0
                .sOrderID = ""
                
                .nCheckOrderNo = 1
                .sOrderNO = txtOrder.Text
            Else
                .nCheckOrderID = 1
                .sOrderID = txtOrder.Text
                
                .nCheckOrderNo = 0
                .sOrderNO = ""
            End If
        Else
            .nCheckOrderID = 0
            .sOrderID = ""
            .nCheckOrderNo = 0
            .sOrderNO = ""
        End If
        
        
        .sDate = MakeDate(DF_SHORT, dtpDate(0))
        .eDate = MakeDate(DF_SHORT, dtpDate(1))
        
        .nCheckCutom = IIf(chkCustom.Value = vbChecked, 1, 0)
        .sCustomID = Trim(txtCustomID.Tag)
        
        .nCheckCardID = IIf(chkCardSearch.Value = vbChecked, 1, 0)
        .sCardID = txtCardID.Text
        
        
        .nCheckTeam = IIf(chkTeamID.Value = vbChecked, 1, 0)
        .sTeamID = sTeamID
        
        
        .nCheckArticle = IIf(chkArticleID.Value = vbChecked, 1, 0)
        .sArticleID = Trim(txtArticleID.Tag)

    End With
    
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_MatchColor_sView"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TParaValue.sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TParaValue.eDate)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaValue.nCheckOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TParaValue.sOrderID)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaValue.nCheckOrderNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TParaValue.sOrderNO)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaValue.nCheckCutom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaValue.sCustomID)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaValue.nCheckCardID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TParaValue.sCardID)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaValue.nCheckTeam)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TParaValue.sTeamID)
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, TParaValue.nCheckArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TParaValue.sArticleID)
    End With
    Set rs = adoCmd.Execute
    Set adoCmd = Nothing
    
    With grdData
        .Rows = .FixedRows
        .Redraw = flexRDNone
        .ExplorerBar = flexExNone

        Do Until rs.EOF
            
            .AddItem "" & vbTab & IIf(rs!ReWorkClss = "*", "■", "") & vbTab & MakeDate(DF_MD, rs!ResultDate) & vbTab & MakeCardID(rs!CardID, OM_EXPAND) & vbTab & Trim(rs!kCustom) & vbTab & _
                        Trim(rs!ArticleName) & vbTab & Trim(rs!OrderNo) & vbTab & MakeOrderID(rs!OrderID, OM_COMPACT) & vbTab & Trim(rs!ColorName) & vbTab & _
                        rs!Roll & vbTab & Format(rs!Qty, "#,##0") & vbTab & rs!Custom & vbTab & rs!ReWorkReason & vbTab & rs!TeamName & vbTab & _
                        rs!PersonName & vbTab & MakeDate(DF_MD, rs!SetDate) & vbTab & rs!WorkSeq & vbTab & SetCurrency(rs!StuffWidth, 2) & vbTab & rs!StuffDensity
            .RowHeight(.Rows - 1) = 350
            
            dTotRoll = dTotRoll + rs!Roll
            dTotQty = dTotQty + rs!Qty
            
            If rs!ReWorkClss = "*" Then
                dReWorkRoll = dReWorkRoll + rs!Roll
                dReWorkQty = dReWorkQty + rs!Qty
            End If
            
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        .Redraw = flexRDDirect
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        Else
            MsgBox LoadResString(203), vbInformation
        End If
    End With
    
    With grdSum
        .Rows = 0
        .AddItem ""
        .TextArray(0) = "총  배  색  량"
        .TextArray(1) = Format(dTotRoll, "#,##0 절")
        .TextArray(2) = Format(dTotQty, "#,##0 YDS")
        .Cell(flexcpFontSize, 0, 1, 0, 2) = 12
        .Cell(flexcpFontBold, 0, 1, 0, 2) = True
    
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 0) = "배  색  량"
        .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 1) = SetCurrency(dTotRoll - dReWorkRoll, 0) & " 절"
        
        .Cell(flexcpText, .Rows - 1, 2, .Rows - 1, 2) = SetCurrency(dTotQty - dReWorkQty, 0) & " YDS"
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        .MergeRow(.Rows - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, 2) = 12
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 0) = "수 정  배  색  량"
        .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 1) = SetCurrency(dReWorkRoll, 0) & " 절"
        
        .Cell(flexcpText, .Rows - 1, 2, .Rows - 1, 2) = SetCurrency(dReWorkQty, 0) & " YDS"
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, 2) = 12
        .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
    
    End With
    
    With grdData
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 7) = "총  배  색  량"
        .Cell(flexcpText, .Rows - 1, 8, .Rows - 1, 11) = SetCurrency(dTotRoll, 0) & " 절"
        
        .Cell(flexcpText, .Rows - 1, 12, .Rows - 1, .Cols - 1) = SetCurrency(dTotQty, 0) & " YDS"
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        .MergeRow(.Rows - 1) = True
        .RowHidden(.Rows - 1) = True
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 7) = "배  색  량"
        .Cell(flexcpText, .Rows - 1, 8, .Rows - 1, 11) = SetCurrency(dTotRoll - dReWorkRoll, 0) & " 절"
        
        .Cell(flexcpText, .Rows - 1, 12, .Rows - 1, .Cols - 1) = SetCurrency(dTotQty - dReWorkQty, 0) & " YDS"
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        .MergeRow(.Rows - 1) = True
        .RowHidden(.Rows - 1) = True
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 7) = "수 정  배  색  량"
        .Cell(flexcpText, .Rows - 1, 8, .Rows - 1, 11) = SetCurrency(dReWorkRoll, 0) & " 절"
        
        .Cell(flexcpText, .Rows - 1, 12, .Rows - 1, .Cols - 1) = SetCurrency(dReWorkQty, 0) & " YDS"
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
        .MergeRow(.Rows - 1) = True
        .RowHidden(.Rows - 1) = True

    End With
    
End Sub


Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
    Case 3
        Call ReturnCode(LG_ORDER, , False, txtOrder)
    Case 4
        Call ReturnCode(LG_CUSTOM, 0, False, txtCustomID)
    Case 5
        Call ReturnCode(LG_ARTICLE, , False, txtArticleID)
    End Select
End Sub



Private Sub Form_Load()
    Dim i%
    Dim dSql_str$, dRS As New ADODB.Recordset
    
    Me.Move 0, 0, 15300, 9660

    Call SetOperate(Me)
    
    dtpDate(0) = Now
    dtpDate(1) = Now
    
    
    dSql_str = "SELECT Team FROM [mt_team] WHERE UseClss = '' ORDER BY TeamID "
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    If dRS.RecordCount > 0 Then
        Call FillComboBox(cboTeamID, dRS)
    End If
    dRS.Close
    Set dRS = Nothing



    For i = 3 To 5
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        cmdFind(i).Enabled = False
        cboTeamID.Enabled = False
    Next i
    
    Call InitGrid
    
    i = ModifyGrid
    
    Show

End Sub

Private Sub cmdSearch_Click()
    grdData.Rows = grdData.FixedRows
    grdDetail.Rows = grdDetail.FixedRows
    
    With grdSum
        .TextArray(1) = Format(0, "#,##0 절")
        .TextArray(2) = Format(0, "#,##0 YDS")
        .Cell(flexcpFontSize, 0, 1, 0, 2) = 12
        .Cell(flexcpFontBold, 0, 1, 0, 2) = True
    End With
    
    Call FillGridData

End Sub



Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub



Private Sub grdData_Click()
    With grdData
        If .Rows = .FixedRows Then Exit Sub
            
        Call FillMatchViewDetail(Replace(.TextMatrix(.Row, 3), "-", ""), .TextMatrix(.Row, .Cols - 3))
    End With
End Sub

Private Sub optOrder_Click(Index As Integer)
    chkOrder.Caption = optOrder(Index).Caption
End Sub


Private Sub FillMatchViewDetail(ByVal CardID As String, ByVal WorkSeq As Integer)
    Dim adoCmd As ADODB.Command
    Dim rsData As New ADODB.Recordset
    Dim nRows%, nCols%

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_MatchColor_sViewDetail"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, CardID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, WorkSeq)
        
    End With
    Set rsData = adoCmd.Execute
    Set adoCmd = Nothing

    
    With grdDetail
        .Redraw = flexRDNone
        .Rows = .FixedRows
        Do Until rsData.EOF
            '---맨처음 레코드 넣기
            
            If (.Rows = .FixedRows Or Trim(.TextMatrix(nRows, 0)) <> CheckNull(rsData!RollGroup) Or nCols >= 16) Then
                .Rows = .Rows + 1
                nRows = .Rows - 1
                nCols = 0       '--- RollGroup 설정
                .TextMatrix(nRows, nCols) = CheckNull(rsData!RollGroup)
               ' .ff.FixedAlignment = flexAlignRightCenter
                nCols = nCols + 1
            End If
            
            
            
''            If .Rows = .FixedRows Then
''                .Rows = .Rows + 1
''                nRows = .Rows - 1
''                nCols = 0       '--- RollGroup 설정
''                .TextMatrix(nRows, nCols) = CheckNull(rsData!RollGroup)
''                .FixedAlignment = flexAlignRightCenter
''                nCols = nCols + 1
''            ElseIf Trim(.TextMatrix(nRows, 0)) <> CheckNull(rsData!RollGroup) Or nCols >= 16 Then
''                    .Rows = .Rows + 1
''                    nRows = .Rows - 1
''                    nCols = 0
''                    .TextMatrix(nRows, nCols) = CheckNull(rsData!RollGroup)
''                    .FixedAlignment = flexAlignRightCenter
''                    nCols = nCols + 1
''            End If
            .TextMatrix(nRows, nCols) = rsData!RollQty
            nCols = nCols + 1
            rsData.MoveNext
        Loop
        rsData.Close
        Set rsData = Nothing
            
        Dim iCount As Integer
        For iCount = .FixedRows To .Rows - 1
            .RowHeight(iCount) = 350
        Next iCount
        
        .Redraw = flexRDDirect
    End With
    
End Sub



Private Sub cmdExit_Click()
    PlusMDI.pnlMenu.Visible = True
    Unload Me
End Sub

Private Sub InitGrid()
    Dim iCount As Integer
    Dim i%
    
    '---- RollGroup의 Detail
    Call SetVSFlexGrid(grdDetail)
    With grdDetail
        .Redraw = flexRDNone
        .WordWrap = False
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 16
        .Rows = .FixedRows + 1
        .TextArray(0) = "":         .ColAlignment(0) = flexAlignCenterCenter
        .ColWidth(0) = LIMIT_WIDTH3
        
        
        For i = 1 To .Cols - 1
            .TextArray(i) = i
            .ColWidth(i) = Int((.Width - .ColWidth(0)) / 15)
            .ColAlignment(iCount) = flexAlignRightCenter
        Next i
        
        .TextMatrix(0, 0) = "그  룹": .ColAlignment(0) = flexAlignCenterCenter
        
        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusHeavy
        .Redraw = flexRDBuffered
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect

    End With
    
    With grdSum
    
        .Redraw = flexRDNone
        .WordWrap = False
        .Rows = 1
        .FixedRows = 0
        .Cols = 3
        .FixedCols = 1
        
        .RowHeight(0) = 350
        .ColWidth(0) = 5000

        .ScrollBars = flexScrollBarNone
        .HighLight = flexHighlightNever
        .FillStyle = flexFillRepeat
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False

        .RowHeightMin = 275
        .WordWrap = True
        .ExtendLastCol = True
        
        .ColAlignment(0) = flexAlignCenterCenter
        
        For iCount = 0 To .Cols - 1
            .FixedAlignment(iCount) = flexAlignCenterCenter
        Next iCount

        .Redraw = True
        
        .TextArray(0) = "합계"
        .TextArray(1) = "0 건":         .ColWidth(1) = 7000
        .TextArray(2) = "0 YDS"
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        
        .Redraw = flexRDDirect
    End With
End Sub




Private Function ModifyGrid() As Integer
    Dim i%
    Dim nProcess As EPROCESSCODE
    
    Call SetVSFlexGrid(grdData)
    
    With grdData
        .Cols = 19
        .Rows = 4
        .FixedRows = 4
        .ScrollBars = flexScrollBarBoth
        .WordWrap = False
        
        .RowHeightMin = 300
        .RowHeight(3) = 400
        
        
        .TextMatrix(3, 0) = "":                         .ColWidth(0) = 0
        .TextMatrix(3, 1) = "구분":                     .ColWidth(1) = 400:             .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = "실적일":                   .ColWidth(2) = 650:             .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(3, 3) = "카드번호":                 .ColWidth(3) = 1200:            .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(3, 4) = "거래처":                   .ColWidth(4) = 1400:            .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(3, 5) = "품명":                     .ColWidth(5) = 2200:            .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(3, 6) = "OrderNo":                  .ColWidth(6) = 1200:            .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(3, 7) = "관리번호":                 .ColWidth(7) = 800:             .ColAlignment(7) = flexAlignCenterCenter
        .TextMatrix(3, 8) = "색상명":                   .ColWidth(8) = 1500:            .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(3, 9) = "절수":                     .ColWidth(9) = 600:             .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(3, 10) = "수량":                    .ColWidth(10) = 800:            .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(3, 11) = "제직처":                  .ColWidth(11) = 1000:           .ColAlignment(11) = flexAlignCenterCenter
        .TextMatrix(3, 12) = "수정원인":                .ColWidth(12) = 1200:           .ColAlignment(12) = flexAlignLeftCenter
        .TextMatrix(3, 13) = "작업조":                  .ColWidth(13) = 0:              .ColAlignment(13) = flexAlignCenterCenter
        .TextMatrix(3, 14) = "작업자":                  .ColWidth(14) = 700:            .ColAlignment(14) = flexAlignCenterCenter
        .TextMatrix(3, 15) = "작업일":                  .ColWidth(15) = 700:            .ColAlignment(15) = flexAlignCenterCenter
        .TextMatrix(3, 16) = "workseq":                 .ColWidth(16) = 0:              .ColAlignment(16) = flexAlignCenterCenter
        .TextMatrix(3, 17) = "생지폭":                  .ColWidth(17) = 800:            .ColAlignment(17) = flexAlignRightCenter
        .TextMatrix(3, 18) = "생지밀도":                .ColWidth(18) = 800:            .ColAlignment(18) = flexAlignCenterCenter
        
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .MergeCells = flexMergeFree
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        For i = 0 To 3
            .MergeRow(i) = True
        Next i
        .FixedCols = 0
        
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusSolid
        
        .Redraw = flexRDDirect
    End With

End Function


Private Sub txtCustomID_GotFocus()
    Call GotFocusText(txtCustomID)

End Sub

Private Sub txtOrder_GotFocus()
    Call GotFocusText(txtOrder)

End Sub
