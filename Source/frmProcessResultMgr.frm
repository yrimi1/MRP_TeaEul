VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProcessResultMgr 
   ClientHeight    =   9270
   ClientLeft      =   2055
   ClientTop       =   2595
   ClientWidth     =   15180
   Icon            =   "frmProcessResultMgr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15180
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   780
      Left            =   14190
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   12
      ToolTipText     =   "자료 저장"
      Top             =   30
      Width           =   870
   End
   Begin Threed.SSCommand cmdHTML 
      Height          =   690
      Left            =   8445
      TabIndex        =   16
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      HTML(&H)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   10125
      TabIndex        =   15
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11820
      TabIndex        =   13
      Top             =   8520
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13500
      TabIndex        =   14
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdSum 
      Height          =   420
      Left            =   120
      TabIndex        =   17
      Top             =   7020
      Visible         =   0   'False
      Width           =   15165
      _cx             =   26749
      _cy             =   741
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
      Height          =   7605
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   15180
      _cx             =   26776
      _cy             =   13414
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
   Begin Threed.SSFrame frmSearch 
      Height          =   825
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   1455
      _Version        =   196609
      Begin VB.TextBox txtCardID 
         Height          =   300
         Index           =   1
         Left            =   12240
         TabIndex        =   11
         Top             =   465
         Width           =   1515
      End
      Begin VB.TextBox txtCardID 
         Height          =   300
         Index           =   0
         Left            =   10350
         TabIndex        =   10
         Top             =   465
         Width           =   1425
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1455
         MousePointer    =   99  '사용자 정의
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   465
         Width           =   600
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   1455
         MousePointer    =   99  '사용자 정의
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   75
         Width           =   600
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         Left            =   6630
         TabIndex        =   4
         Top             =   75
         Width           =   1905
      End
      Begin VB.TextBox txtArticle 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   6630
         TabIndex        =   6
         Top             =   465
         Width           =   1905
      End
      Begin VB.TextBox txtOrderID 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   10350
         TabIndex        =   8
         Top             =   90
         Width           =   1905
      End
      Begin Threed.SSPanel pnlOrder 
         Height          =   765
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1349
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   450
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   120
            Width           =   1200
         End
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   3450
         TabIndex        =   1
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   116785153
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   3435
         TabIndex        =   2
         Top             =   465
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   116785153
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   2160
         TabIndex        =   25
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
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
            Caption         =   "배색일자"
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   0
            Top             =   45
            Value           =   1  '확인
            Width           =   1080
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   5280
         TabIndex        =   26
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
            TabIndex        =   3
            Top             =   45
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   8550
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   465
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
         Left            =   5280
         TabIndex        =   28
         Top             =   465
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
            TabIndex        =   5
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   12270
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   90
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
         Left            =   9060
         TabIndex        =   30
         Top             =   90
         Width           =   1260
         _ExtentX        =   2223
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
            TabIndex        =   7
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   0
         Left            =   8550
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   90
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
         Index           =   4
         Left            =   9060
         TabIndex        =   34
         Top             =   465
         Width           =   1260
         _ExtentX        =   2223
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
            TabIndex        =   9
            Top             =   60
            Width           =   1185
         End
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "~"
         Height          =   180
         Index           =   3
         Left            =   11895
         TabIndex        =   35
         Top             =   510
         Width           =   150
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   4755
         TabIndex        =   32
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   4755
         TabIndex        =   31
         Top             =   135
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmProcessResultMgr"
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

Private nChkDate As Integer, sDate As String, eDate As String
Private nChkOrder As Integer, sOrderID As String
Private nChkCustom As Integer, sCustomID As String
Private nChkArticle As Integer, sArticleID As String
Private nChkCard As Integer, sFromCardID  As String, sToCardID  As String

Private Sub cmdPrint_Click()
    Dim nFontSize As Integer
    Dim nColor As Long
    
'    If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbNo Then
'        Exit Sub
'    End If
    
    Screen.MousePointer = vbHourglass
    
    Call ColResize(grdData, ES_REDUCE, 20)
    
    With grdData
        .Redraw = flexRDBuffered
    
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = vbWhite
    
        .RowHeight(0) = 500
        .RowHeight(1) = 400
        .RowHeight(2) = 400
        .RowHeight(3) = 400
        
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        .RowHidden(3) = False
        
        ' Header Tilte
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = " "
        .Cell(flexcpText, 1, 1, 1, .Cols - 1) = "공정관리일지"
        .Cell(flexcpFontSize, 1, 0, 1, .Cols - 1) = 16
        .Cell(flexcpFontBold, 1, 0, 1, .Cols - 1) = True
        
   '     .Cell(flexcpFontUnderline, 1, 1, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 2, 1, 2, 3) = "▶ 수주일자 : " & IIf(nChkDate = 1, Format(dtpDate(0), "YYYY/MM/DD") & " ~ " & Format(dtpDate(1), "YYYY/MM/DD"), AllStr)
        .Cell(flexcpText, 2, 4, 2, 6) = "▶ 거래처   : " & IIf(nChkCustom = 1, Trim(txtCustom), AllStr)
        
        .Cell(flexcpText, 2, 7, 2, .Cols - 1) = "▶ 카드   : " & IIf(nChkCard = 1, Trim(sFromCardID) & " ~ " & Trim(sToCardID), AllStr)
        
        .Cell(flexcpText, 3, 1, 3, 3) = "▶ 품  명   : " & IIf(nChkArticle = 1, txtArticle, AllStr)
        .Cell(flexcpText, 3, 4, 3, 6) = "▶ " & IIf(optOrder(0).Value = True, optOrder(0).Caption, optOrder(1).Caption) & " : " & _
                                                  IIf(nChkOrder = 1, txtCustom, AllStr)

        .Cell(flexcpText, 3, 16, 3, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD HH:SS")
        .Cell(flexcpAlignment, 2, 0, 3, .Cols - 1) = flexAlignLeftCenter
        
        .Cell(flexcpBackColor, 0, 0, 3, .Cols - 1) = vbWhite
'        .SheetBorder = &H80000012
        
        Call SetPrintMode(grdData, 4, True)
        
        .PrintGrid "태을염직", True, 2, 100, 500
        
        Call SetPrintMode(grdData, 4, False)

        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .RowHidden(3) = True
        
        .Redraw = flexRDDirect
    End With
    
    Screen.MousePointer = vbDefault
    
    Call ColResize(grdData, ES_EXPAND, 20)
        
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then       ' 금일
            dtpDate(0) = Date
            dtpDate(1) = Date
    ElseIf Index = 1 Then   ' 금월
            dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
            dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If

End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub


Private Sub cmdExcel_Click()

    If grdData.Rows = 1 Then
        MsgBox LoadResString(203), vbInformation
        Exit Sub
    End If
    Call MakeExcelGrid(grdData)

End Sub

Private Sub cmdFind_Click(Index As Integer)
    Dim dOrderID As String, dOrderNO As String
    
    Select Case Index
        Case 0             '[3] 거래처 코드
            Call ReturnCode(LG_CUSTOM, 0, False, txtCustom)
        Case 1             '[4] 품명
            Call ReturnCode(LG_ARTICLE, , False, txtArticle)
        Case 2             '[4] 품명
            Call ReturnCode(LG_ORDER, , False, txtOrderID)
            dOrderNO = Trim(txtOrderID.Text)
            dOrderID = Trim(txtOrderID.Tag)
            If optOrder(0).Value = True Then   'OrderNO
                txtOrderID.Text = dOrderNO
                txtOrderID.Tag = dOrderID
                
            Else
                txtOrderID.Tag = dOrderNO
                txtOrderID.Text = dOrderID
            End If
            
            
    End Select
End Sub

Private Sub cmdHTML_Click()
    If grdData.Rows = 1 Then
        MsgBox LoadResString(203), vbInformation
        cmdSearch.SetFocus

        Exit Sub
    End If

    If MakeHtmlGrid(grdData, "C:\" & Me.Caption & ".html") Then
        Call RelateOpen(Me.hWnd, "C:\" & Me.Caption & ".html")
    End If

End Sub


Private Sub Form_Load()
    Dim i%
    Me.Move 0, 0, 15300, 9660

    Call SetOperate(Me)
    
    dtpDate(0) = Now
    dtpDate(1) = Now
    

    Call InitGrid
    
    i = ModifyGrid
    
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)    '---거래처
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)    '---품명
    cmdFind(2).Picture = LoadResPicture("FIND", vbResIcon)    '---관리번호
    

    dtpDate(0).Enabled = chkSearch(0).Value
    dtpDate(1).Enabled = chkSearch(0).Value
    txtCustom.Enabled = chkSearch(1).Value
    cmdFind(0).Enabled = chkSearch(1).Value
    txtArticle.Enabled = chkSearch(2).Value
    cmdFind(1).Enabled = chkSearch(2).Value
    txtOrderID.Enabled = chkSearch(3).Value
    cmdFind(2).Enabled = chkSearch(3).Value
    txtCardID(0).Enabled = chkSearch(4).Value
    txtCardID(1).Enabled = chkSearch(4).Value
    
    Show

End Sub


Private Sub cmdSearch_Click()
    Call FillGridData
End Sub



Private Sub chkSearch_Click(Index As Integer)

    Select Case Index
        Case 0
            dtpDate(0).Enabled = chkSearch(0).Value
            dtpDate(1).Enabled = chkSearch(0).Value
            If chkSearch(Index).Value = vbChecked Then
                dtpDate(0).SetFocus
            End If
        Case 1
            txtCustom.Enabled = chkSearch(1).Value
            cmdFind(0).Enabled = chkSearch(1).Value
            If chkSearch(Index).Value = vbChecked Then
                txtCustom.SetFocus
            Else
                txtCustom.Text = ""
                txtCustom.Tag = ""
            End If
        Case 2
            txtArticle.Enabled = chkSearch(2).Value
            cmdFind(1).Enabled = chkSearch(2).Value
            If chkSearch(Index).Value = vbChecked Then
                txtArticle.SetFocus
            Else
                txtArticle.Text = ""
                txtArticle.Tag = ""
            End If
        Case 3
            txtOrderID.Enabled = chkSearch(3).Value
            cmdFind(2).Enabled = chkSearch(3).Value
            If chkSearch(Index).Value = vbChecked Then
                txtOrderID.SetFocus
            Else
                txtOrderID.Text = ""
                txtOrderID.Tag = ""
            End If
        Case 4
            txtCardID(0).Enabled = chkSearch(4).Value
            txtCardID(1).Enabled = chkSearch(4).Value
            If chkSearch(Index).Value = vbChecked Then
                txtCardID(0).SetFocus
            Else
                txtCardID(0).Text = ""
                txtCardID(1).Text = ""
            End If
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub



Private Sub optOrder_Click(Index As Integer)
    
    chkSearch(3).Caption = optOrder(Index).Caption
    Call SetToggle
End Sub


Sub FillGridData()
    Dim oClss As PlusLib2.CProcess
    Dim rs As ADODB.Recordset
    Dim sCardID As String
    
    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass

    Set oClss = New PlusLib2.CProcess
    oClss.Connection = g_adoCon
    oClss.UserName = g_sUserName
    
    nChkDate = 0: sDate = "": eDate = ""
    nChkOrder = 0: sOrderID = ""
    nChkCustom = 0: sCustomID = ""
    nChkArticle = 0: sArticleID = ""
    nChkCard = 0: sFromCardID = "": sToCardID = ""
    
    '수주일자
    If chkSearch(0).Value Then
        nChkDate = 1
        sDate = MakeDate(DF_SHORT, dtpDate(0))
        eDate = MakeDate(DF_SHORT, dtpDate(1))
    End If
    
    '거래처
    If chkSearch(1).Value Then
        nChkCustom = 1
        sCustomID = txtCustom.Tag
    End If
    
    '품명
    If chkSearch(2).Value Then
        nChkArticle = 1
        sArticleID = txtArticle.Tag
    End If
    
    'OrderID, OrderNO
    If chkSearch(3).Value Then
        nChkOrder = 1
        If optOrder(0).Value Then
            sOrderID = txtOrderID.Tag
        Else
            sOrderID = txtOrderID.Text
        End If
    End If
    
    If chkSearch(4).Value = vbChecked Then
        nChkCard = 1
        sFromCardID = txtCardID(0).Text
        sToCardID = txtCardID(1).Text
    End If
    
    Set rs = oClss.GetProcessResultMgr(nChkDate, sDate, eDate, nChkOrder, sOrderID _
                                , nChkCustom, sCustomID, nChkArticle, sArticleID _
                                , nChkCard, sFromCardID, sToCardID)

    Set oClss = Nothing
    
    With grdData
        .Rows = .FixedRows
        .Redraw = flexRDDirect

        Do Until rs.EOF
            sCardID = CStr(Int(Right(rs!CardID, 4))) & IIf(Trim(rs!SplitID) = "", "", "(" & Trim(rs!SplitID) & ")")
            
            .AddItem CStr(.Rows - .FixedRows + 1) & vbTab & Trim(sCardID) & vbTab & Trim(rs!kCustom) & vbTab & _
                     Trim(rs!Article) & vbTab & MakeOrderID(rs!OrderID, OM_COMPACT) & vbTab & Trim(rs!OrderNo) & vbTab & _
                     Trim(rs!Color) & vbTab & rs!Roll & vbTab & SetCurrency(rs!Qty, 0) & vbTab & _
                     IIf(Trim(rs!준비) = "", "", MakeDate(DF_MD, rs!준비)) & vbTab & _
                     IIf(Trim(rs!정련) = "", "", MakeDate(DF_MD, rs!정련)) & vbTab & _
                     IIf(Trim(rs!피치) = "", "", MakeDate(DF_MD, rs!피치)) & vbTab & _
                     IIf(Trim(rs!C염색) = "", "", MakeDate(DF_MD, rs!C염색)) & vbTab & _
                     IIf(Trim(rs!R염색) = "", "", MakeDate(DF_MD, rs!R염색)) & vbTab & _
                     IIf(Trim(rs!DRY) = "", "", MakeDate(DF_MD, rs!DRY)) & vbTab & _
                     IIf(Trim(rs!가공) = "", "", MakeDate(DF_MD, rs!가공)) & vbTab & _
                     IIf(Trim(rs!검사) = "", "", MakeDate(DF_MD, rs!검사)) & vbTab & _
                      rs!GOOD & vbTab & rs!NG & ""
            rs.MoveNext
        Loop
        Screen.MousePointer = vbDefault
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If
        
        .Redraw = flexRDDirect
        
    End With
    
    rs.Close
    Set rs = Nothing
    
    If grdData.Rows > grdData.FixedRows Then
        cmdHTML.Visible = True
        cmdExcel.Visible = True
        cmdPrint.Visible = True
    Else
        cmdHTML.Visible = False
        cmdExcel.Visible = False
        cmdPrint.Visible = False
    End If
    
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, "frmProcessResultMgr.FillGridData", Err.Description)
    Set rs = Nothing
    Set oClss = Nothing

End Sub

Private Sub cmdExit_Click()
    PlusMDI.pnlMenu.Visible = True
    Unload Me
End Sub

Private Sub InitGrid()
    Dim iCount As Integer
    
    With grdSum
    
        .Redraw = flexRDNone
        
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
        .WordWrap = False
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
    Dim i%, iHeadRow As Integer
    Dim nProcess As EPROCESSCODE
    
    Call SetVSFlexGrid(grdData)
    
    With grdData
        .Cols = 21
        .Rows = 6
        .FixedRows = 6
        
        ' 0~2번 Row는 리포트 발행시 타이틀및 일자등 출력하는 부분
        ' 3,4번 Row는 실제 화면에서 컬럼명 출력부분
        
        For i = 0 To 4
            .RowHeight(i) = 300
        Next i
        
        .RowHeight(4) = 400
        .RowHeightMin = 300
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .RowHidden(3) = True
        
        iHeadRow = 4
        
        ' 기본내역
        .TextMatrix(iHeadRow, 0) = " ":                        .ColWidth(0) = 300
        .TextMatrix(iHeadRow, 1) = "밧자" & vbCrLf & "번호":   .ColWidth(1) = 800:             .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(iHeadRow, 2) = "거래처명 ":                .ColWidth(2) = 1200:            .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 3) = "품명 ":                    .ColWidth(3) = 2000:            .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 4) = "관리번호":                 .ColWidth(4) = 800:             .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 5) = "OrderNO":                  .ColWidth(5) = 1300:            .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 6) = "색상명":                   .ColWidth(6) = 1600:            .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(iHeadRow, 7) = "절수":                     .ColWidth(7) = 500:             .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(iHeadRow, 8) = "수량":                     .ColWidth(8) = 800:             .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(iHeadRow, 9) = "준비":                     .ColWidth(9) = 600:             .ColAlignment(9) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 10) = "정련":                    .ColWidth(10) = 600:            .ColAlignment(10) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 11) = "피치":                    .ColWidth(11) = 600:            .ColAlignment(11) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 12) = "CPB":                     .ColWidth(12) = 600:            .ColAlignment(12) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 13) = "염색":                    .ColWidth(13) = 600:            .ColAlignment(13) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 14) = "DRY":                     .ColWidth(14) = 600:            .ColAlignment(14) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 15) = "가공":                    .ColWidth(15) = 600:            .ColAlignment(15) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 16) = "검사":                    .ColWidth(16) = 600:            .ColAlignment(16) = flexAlignCenterCenter
        .TextMatrix(iHeadRow, 17) = "검사수량":                .ColWidth(17) = 600:            .ColAlignment(17) = flexAlignRightCenter
        .TextMatrix(iHeadRow, 18) = "검사수량":                .ColWidth(18) = 600:            .ColAlignment(18) = flexAlignRightCenter
        .TextMatrix(iHeadRow, 19) = "출고" & vbCrLf & "수량":  .ColWidth(19) = 600:            .ColAlignment(19) = flexAlignRightCenter
        .TextMatrix(iHeadRow, 20) = "사고내역":                .ColWidth(20) = 1400:           .ColAlignment(20) = flexAlignCenterCenter
        
        
        iHeadRow = iHeadRow + 1
        
        ' 기본내역
        .TextMatrix(iHeadRow, 0) = " "
        .TextMatrix(iHeadRow, 1) = "밧자" & vbCrLf & "번호"
        .TextMatrix(iHeadRow, 2) = "거래처명 "
        .TextMatrix(iHeadRow, 3) = "품명 "
        .TextMatrix(iHeadRow, 4) = "관리번호"
        .TextMatrix(iHeadRow, 5) = "OrderNO"
        .TextMatrix(iHeadRow, 6) = "색상명"
        .TextMatrix(iHeadRow, 7) = "절수"
        .TextMatrix(iHeadRow, 8) = "수량"
        .TextMatrix(iHeadRow, 9) = "준비"
        .TextMatrix(iHeadRow, 10) = "정련"
        .TextMatrix(iHeadRow, 11) = "피치"
        .TextMatrix(iHeadRow, 12) = "CPB"
        .TextMatrix(iHeadRow, 13) = "염색"
        .TextMatrix(iHeadRow, 14) = "DRY"
        .TextMatrix(iHeadRow, 15) = "가공"
        .TextMatrix(iHeadRow, 16) = "검사"
        .TextMatrix(iHeadRow, 17) = "합격"
        .TextMatrix(iHeadRow, 18) = "불량"
        .TextMatrix(iHeadRow, 19) = "출고" & vbCrLf & "수량":
        .TextMatrix(iHeadRow, 20) = "사고내역"
        
        .ColHidden(11) = True
        .ColHidden(14) = True
        
        
        Call FixedColAlignMentSetting(grdData)
        Dim II%
        For II = 0 To .Rows - 1
            .MergeRow(II) = True
        Next II
        
        For II = 0 To .Cols - 1
            .MergeCol(II) = True
        Next
        
        Call SetToggle
        
        .MergeCells = flexMergeFixedOnly
        .WordWrap = False
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        
        .Redraw = flexRDDirect
    End With
End Function

Sub SetToggle()
    Dim dOrderNO As String, dOrderID As String
    
    If optOrder(0).Value = True Then
        grdData.ColWidth(4) = 0
        grdData.ColWidth(5) = 1300
    Else
        grdData.ColWidth(4) = 800
        grdData.ColWidth(5) = 0
    End If
    
    If chkSearch(3).Value = vbChecked Then
        If optOrder(0).Value = True Then   'OrderNO
            dOrderID = txtOrderID.Text
            dOrderNO = txtOrderID.Tag
            
            txtOrderID.Text = dOrderNO
            txtOrderID.Tag = dOrderID
        Else
            dOrderID = txtOrderID.Tag
            dOrderNO = txtOrderID.Text
            
            txtOrderID.Text = dOrderID
            txtOrderID.Tag = dOrderNO
            
            
        End If
    End If
End Sub


Private Sub txtArticle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click(1)
    End If

End Sub

Private Sub txtCardID_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = vbKeyReturn Then
        txtCardID(1).Text = txtCardID(0)
        Call MoveFocus(KeyCode)
    End If
End Sub

Private Sub txtCustom_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtCustom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click(0)
    End If
End Sub


Private Sub txtOrderID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click(2)
    End If

End Sub
