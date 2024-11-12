VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   ClientHeight    =   6795
   ClientLeft      =   2220
   ClientTop       =   1440
   ClientWidth     =   5850
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmFind.frx":000C
   ScaleHeight     =   6795
   ScaleWidth      =   5850
   Begin MSComctlLib.ProgressBar proProgress 
      Height          =   375
      Left            =   1725
      TabIndex        =   19
      Top             =   2445
      Visible         =   0   'False
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin Threed.SSPanel pnlChoice 
      Align           =   3  '왼쪽 맞춤
      Height          =   6000
      Left            =   0
      TabIndex        =   2
      Top             =   795
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   10583
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   26
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   38
         Top             =   5565
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   25
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   37
         Top             =   5565
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   24
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   36
         Top             =   5190
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   23
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   35
         Top             =   5190
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   22
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   34
         Top             =   4815
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   21
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   33
         Top             =   4815
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   20
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   32
         Top             =   4440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   19
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   31
         Top             =   4440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   18
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   30
         Top             =   4065
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   17
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   29
         Top             =   4065
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   16
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   28
         Top             =   3690
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Height          =   345
         Index           =   15
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   27
         Top             =   3690
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton optChoice 
         Caption         =   "영문"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   375
         Width           =   735
      End
      Begin VB.OptionButton optChoice 
         Caption         =   "한글"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   105
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "파"
         Height          =   345
         Index           =   13
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   17
         Top             =   3315
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "하"
         Height          =   345
         Index           =   14
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   18
         Top             =   3315
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "타"
         Height          =   345
         Index           =   12
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   16
         Top             =   2940
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "카"
         Height          =   345
         Index           =   11
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   15
         Top             =   2940
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "차"
         Height          =   345
         Index           =   10
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   14
         Top             =   2565
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "자"
         Height          =   345
         Index           =   9
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   13
         Top             =   2565
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "아"
         Height          =   345
         Index           =   8
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   12
         Top             =   2190
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "사"
         Height          =   345
         Index           =   7
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   11
         Top             =   2190
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "바"
         Height          =   345
         Index           =   6
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   10
         Top             =   1815
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "마"
         Height          =   345
         Index           =   5
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   9
         Top             =   1815
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "라"
         Height          =   345
         Index           =   4
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   8
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "다"
         Height          =   345
         Index           =   3
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   7
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "나"
         Height          =   345
         Index           =   2
         Left            =   480
         MousePointer    =   99  '사용자 정의
         TabIndex        =   6
         Top             =   1065
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "가"
         Height          =   345
         Index           =   1
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   5
         Top             =   1065
         Width           =   375
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "기타"
         Height          =   345
         Index           =   0
         Left            =   60
         MousePointer    =   99  '사용자 정의
         TabIndex        =   4
         Top             =   690
         Width           =   810
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   15
         X2              =   930
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   15
         X2              =   930
         Y1              =   615
         Y2              =   615
      End
   End
   Begin Threed.SSPanel pnlFind 
      Align           =   1  '위 맞춤
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   1402
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdCancel 
         Caption         =   "취소 (&X)"
         Height          =   495
         Left            =   4740
         MousePointer    =   99  '사용자 정의
         TabIndex        =   40
         Top             =   150
         Width           =   1005
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   990
         TabIndex        =   21
         Top             =   420
         Width           =   1365
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Left            =   990
         TabIndex        =   20
         Top             =   75
         Width           =   1365
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "선택 (&O)"
         Height          =   495
         Left            =   3615
         MousePointer    =   99  '사용자 정의
         TabIndex        =   3
         Top             =   150
         Width           =   1005
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색 (&S)"
         Height          =   495
         Left            =   2505
         MousePointer    =   99  '사용자 정의
         TabIndex        =   1
         Top             =   150
         Width           =   1005
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   0
         Left            =   75
         TabIndex        =   22
         Top             =   75
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "코  드"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   1
         Left            =   75
         TabIndex        =   23
         Top             =   420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "명  칭"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   4965
      Left            =   960
      TabIndex        =   39
      Top             =   825
      Width           =   4845
      _cx             =   8546
      _cy             =   8758
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
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "test"
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
      Left            =   2850
      TabIndex        =   26
      Top             =   5580
      Width           =   375
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL As String

Dim m_sCodeField$, m_sNameField$, m_sOrderField$
Dim m_iLimitWidth%, m_iLimitForm%

Dim m_bSelected As Boolean
Dim wData()
Dim dOrderByStr As String

'********************************************************
'*
'* Description: 한글명 배열, 영문명 배열
'*
'********************************************************
Private m_vEnglish As Variant
Private m_vKorean As Variant

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub optChoice_Click(Index As Integer)
    Dim i%, j%

    If optChoice(0).Value Then  ' 한글
        For i = 0 To UBound(m_vKorean)
            cmdChoice(i).Caption = m_vKorean(i)
        Next i
        For i = UBound(m_vKorean) + 1 To cmdChoice.Count - 1
            cmdChoice(i).Visible = False
        Next i
    Else  ' 영문
        For i = 0 To UBound(m_vEnglish)
            cmdChoice(i).Caption = m_vEnglish(i)
        Next i
        
        For i = UBound(m_vKorean) To cmdChoice.Count - 1
            cmdChoice(i).Visible = True
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i%

    m_vEnglish = Array("ELSE", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                        "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    m_vKorean = Array("기타", "가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하")

    proProgress.Visible = False
    m_iLimitWidth = 2350
    With grdData
        .Redraw = False
        
        .Cols = 3
        .Rows = 1
        .RowHeight(0) = 450
        .RowHeightMin = 290
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

        .WordWrap = True
        
        .TextArray(0) = "순위"
        .TextArray(1) = "코드"
        .TextArray(2) = " 명칭"
        
        .ColWidth(0) = 380
        .ColWidth(1) = 1250
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .Redraw = True
    End With
    proProgress.Visible = False
    
    cmdSearch.MousePointer = ssCustom
    cmdSearch.MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    cmdOK.MousePointer = ssCustom
    cmdOK.MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    For i = 0 To cmdChoice.Count - 1
        cmdChoice(i).MousePointer = ssCustom
        cmdChoice(i).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    Next i
    
    m_iLimitForm = Me.Width
End Sub


Public Function SetMsg(SelData(), ByVal Large As Integer, Optional Middle, Optional NewData) As Boolean
    Dim i%
    
    Dim rs As ADODB.Recordset

    On Error Resume Next
    
    '=================================================================================================='
    
    If Large = 0 Then       '[1] 거래처 코드
        m_sCodeField = "CustomID"
        m_sNameField = "KCustom"
        
        SQL = _
            "SELECT CustomID AS [코드], KCustom AS [상호] " & _
            "FROM [mt_Custom] " & _
            "WHERE UseClss = '' "
            
        dOrderByStr = " ORDER BY [코드] "
            
        If Not IsMissing(Middle) Then
            SQL = SQL & "AND TradeID = '" & Format(Middle + 1, "0") & "' "
        End If
    '--------------------------------------------------------------------------------------------------'
    ElseIf Large = 1 Then   '[2] 품명 코드
        m_sCodeField = "ArticleID"
        m_sNameField = "Article"
        
        SQL = _
            "SELECT ArticleID AS [코드], Article AS [품명] " & _
            "FROM [mt_Article] " & _
            "WHERE UseClss = '' "
        dOrderByStr = " ORDER BY [코드] "
    '--------------------------------------------------------------------------------------------------'
    ElseIf Large = 2 Then   '[3] 사원 코드
        m_sCodeField = "A.PersonID"
        m_sNameField = "A.Name"
    
        SQL = _
            "SELECT A.PersonID AS [코드], A.Name AS [성명], B.Depart AS [부서] " & _
            "FROM [mt_Person] A, [mt_Depart] B " & _
            "WHERE A.DepartID = B.DepartID AND A.UseClss = '' "
        
        dOrderByStr = " ORDER BY [코드] "
        If Not IsMissing(Middle) Then
            SQL = SQL & "AND A.DepartID = " & CStr(Middle) & " "
        End If
    '--------------------------------------------------------------------------------------------------'
    ElseIf Large = 3 Then   '[4] 불량 코드
        m_sCodeField = "DefectID"
        m_sNameField = "KDefect"
    
        SQL = _
            "SELECT DefectID AS [코드], KDefect AS [불량명], TagName AS [Tag 명] " & _
            "FROM [mt_Defect] "
        If Not IsMissing(Middle) Then
            SQL = SQL & "WHERE DefectID like '" & CStr(Middle) & "%' "
        End If
        dOrderByStr = " ORDER BY [코드] "
    '--------------------------------------------------------------------------------------------------'
    ElseIf Large = 4 Then   '[5] 수주
        m_sCodeField = "A.OrderID"
        m_sNameField = "A.OrderNo"
        
        SQL = _
            "SELECT A.OrderID AS [ 관리번호 ], A.OrderNo AS [Order No], B.KCustom AS [거래처], " & _
                    "C.Article AS [품명], D.WorkName AS [가공], " & _
                    "E.StuffWidth AS [가공폭], " & _
                    "A.DvlyDate AS [납기], A.OrderQty AS [수주량], " & _
                    "A.CustomID AS [거래처코드], A.ArticleID AS [품명코드], A.WorkID as [가공구분코드], A.UnitClss AS [단위], " & _
                    "(SELECT MAX(UnitPrice) FROM [OrderColor] WHERE OrderID = A.OrderID) AS [단가] " & _
            "FROM [Order] A, [mt_Custom] B, [mt_Article] C, [mt_Work] D, [mt_StuffWidth] E " & _
            "WHERE A.CustomID = B.CustomID " & _
                "AND A.ArticleID = C.ArticleID " & _
                "AND A.WorkID = D.WorkID " & _
                "AND A.WorkWidth = E.StuffWidthID "
                
'                "AND A.CloseClss= '' "
        dOrderByStr = " ORDER BY [ 관리번호 ] "
    '--------------------------------------------------------------------------------------------------'
    ElseIf Large = 5 Then   ' 염료 코드
        m_sCodeField = "DyeAuxID"
        m_sNameField = "DyeAux"
    
        SQL = _
            "SELECT DyeAuxID AS [코드], DyeAux AS [염료명] " & _
            "FROM [mt_DyeAux] " & _
            "WHERE DyeAuxID LIKE '1%' AND ISNULL(UseClss, '') NOT IN ('*') "
        dOrderByStr = " ORDER BY [코드] "
    '--------------------------------------------------------------------------------------------------'
    ElseIf Large = 6 Then   ' 조제 코드
        m_sCodeField = "DyeAuxID"
        m_sNameField = "DyeAux"

        SQL = _
            "SELECT DyeAuxID AS [코드], DyeAux AS [염료명] " & _
            "FROM [mt_DyeAux] " & _
            "WHERE DyeAuxID LIKE '0%' AND ISNULL(UseClss, '') NOT IN ('*') "
        dOrderByStr = " ORDER BY [코드] "
            
    ElseIf Large = 7 Then ' 가공코드
        m_sCodeField = "WorkID"
        m_sNameField = "Work"

        SQL = "SELECT WorkID AS [코드], WorkName AS [가공명] " & _
              "FROM [mt_Work] " & _
              "WHERE UseClss = '' "
        dOrderByStr = " ORDER BY [코드] "
              
    ElseIf Large = 8 Then   ' 사종구분
        m_sCodeField = "ThreadID"
        m_sNameField = "Thread"
        
        SQL = "SELECT ThreadID AS [코드], Thread AS [사종명] " & _
              "FROM [mt_Thread] " & _
              "WHERE UseClss = '' "
        dOrderByStr = " ORDER BY [코드] "
        
    ElseIf Large = 9 Then   ' 원단폭
        m_sCodeField = "StuffWidthID"
        m_sNameField = "StuffWidth"
        
        SQL = "SELECT StuffWidthID AS [코드], StuffWidth AS [원단폭] " & _
              "FROM [mt_StuffWidth] " & _
              "WHERE UseClss = '' "
        dOrderByStr = " ORDER BY [코드] "
              
    ElseIf Large = 10 Then   ' 공정명
        m_sCodeField = "ProcessID"
        m_sNameField = "Process"
        
        SQL = "SELECT ProcessID As [코드], Process AS [공정명]" & _
              "FROM [mt_Process] WHERE ProcessID NOT LIKE '%00' "
        dOrderByStr = " ORDER BY [코드] "
    
                
    End If
    '=================================================================================================='

    If IsMissing(NewData) Then  ' 찾고자하는 데이타가 없을 경우 빈 폼 뛰우기
        Me.Show vbModal
    Else
        Call SetGrid(FL_BY_NAME, NewData) ' [1] 명칭으로 찾기
        If grdData.Rows = 1 Then
            txtName = NewData
            Call SetGrid(FL_BY_CODE, NewData) ' [2] 코드로 찾기 (코드검색이 않되었을 경우)
        End If
        
        With grdData
            If .Rows > .FixedRows Then
                If .Rows = .FixedRows + 1 Then
                    Call SelectData
                Else
                    Me.Show vbModal
                End If
            End If
        End With
    End If
    
    '=================================================================================================='
    If m_bSelected Then
        With grdData
            ReDim SelData(.Cols - 1)
            For i = 0 To .Cols - 1
                SelData(i) = wData(i)
            Next i
        End With
    End If
    
    SetMsg = m_bSelected
End Function

Private Sub SetGrid(ByVal Index As EFindClss, ByVal NewData As String)
    Dim Query$, iGridWidth%, i%
    Dim rs As ADODB.Recordset
    
    Dim sTemp As String
    
    Screen.MousePointer = vbHourglass
    
    If InStr(SQL, "WHERE") Then
        Query = SQL & "AND ("
    Else
        Query = SQL & "WHERE "
    End If
    
    '----------------------------------------------------------------------------------------------'
    If Index = FL_BY_CODE Then        '[1] 코드로 찾기
        Query = Query & m_sCodeField & " = '" & NewData & "') "
    '----------------------------------------------------------------------------------------------'
    ElseIf Index = FL_BY_NAME Then    '[2] 명칭으로 찾기
        If InStr(SQL, "WHERE") Then
            Query = Query & m_sNameField & " LIKE '%" & NewData & "%' OR " & m_sNameField & " LIKE '(주)" & NewData & "%') "
        Else
            Query = Query & m_sNameField & " LIKE '%" & NewData & "%' OR " & m_sNameField & " LIKE '(주)" & NewData & "%' "
        End If
    '----------------------------------------------------------------------------------------------'
    ElseIf Index = FL_BY_BTN Then     '[3] 버튼으로 선택하여 찾기
        i = CInt(NewData)
        
        If optChoice(0).Value Then ' 한글일 경우
            If i = 0 Or i = 14 Then ' [기타] 또는 [하]
                sTemp = "힝힝힝"
            Else
                sTemp = cmdChoice(i + 1).Caption
            End If
        Else ' 영문일 경우
            If i = 0 Then  ' [ELSE]
                sTemp = "zzzz"
            ElseIf i = 26 Then ' [Z]
                sTemp = "ZZZZ"
            Else
                sTemp = cmdChoice(i + 1).Caption
            End If
        End If
        
        If i = 0 Then
            Query = Query & m_sNameField & " < '" & cmdChoice(1).Caption & "' OR " & m_sNameField & " > '" & sTemp & "') "
        Else
            Query = Query & "((" & m_sNameField & " >= '" & cmdChoice(i).Caption & "' " & _
                       "AND " & m_sNameField & " < '" & sTemp & "') "
            If optChoice(1).Value Then ' 영문일때
                Query = Query & "OR (" & m_sNameField & " >= '" & Chr(Asc(cmdChoice(i).Caption) + 32) & "' " & _
                       "AND " & m_sNameField & " < '" & LCase(sTemp) & "'))) "
            Else ' 한글일때
                 Query = Query & "OR (" & m_sNameField & " >= '(주)" & cmdChoice(i).Caption & "' " & _
                        "AND " & m_sNameField & " < '(주)" & sTemp & "'))) "

            End If
        End If
    End If
    '=============================================================================================='

    lblCount = "0"
    
    Query = Query & dOrderByStr
    
    Set rs = New ADODB.Recordset
    rs.Open Query, adoCon, adOpenStatic, adLockReadOnly
    
    With grdData
        .Redraw = False
        
        .RowHeight(0) = 520
        .Rows = rs.RecordCount + 1
        .Cols = rs.Fields.Count
        .ColAlignment(0) = flexAlignCenterCenter
     
        If .Rows > 50 Then
            proProgress.Visible = True
            proProgress.Value = 0
        End If
     
        'Resultset이 가지는 칼럼의 이름을 그리드의 항목명으로 사용한다.
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        
            Select Case rs(i).Type
                Case adInteger, adNumeric, adDecimal, adDouble, adSingle
                    .ColWidth(i) = 1250
                    .ColAlignment(i) = flexAlignRightCenter
                Case adChar
                    .ColWidth(i) = Max(TextWidth(String(rs(i).DefinedSize, "r")), TextWidth(rs(i).Name) + 250)
                    .ColAlignment(i) = flexAlignCenterCenter
                Case adVarChar
                    .ColWidth(i) = Max(TextWidth(String(rs(i).DefinedSize, "v")), TextWidth(rs(i).Name) + 250)
                    .ColAlignment(i) = flexAlignLeftCenter
                Case adVarWChar
                    .ColWidth(i) = Max(TextWidth(String(rs(i).DefinedSize, "i")), TextWidth(rs(i).Name) + 250)
                    .ColAlignment(i) = flexAlignLeftCenter
                Case Else
                    .ColWidth(i) = Max(TextWidth(String(rs(i).DefinedSize, "A")), TextWidth(rs(i).Name) + 250)
            End Select
            iGridWidth = iGridWidth + .ColWidth(i)
        Next i
     
        If iGridWidth > 10000 Then
            iGridWidth = 10000
            .ScrollBars = flexScrollBarBoth
        Else
            .ScrollBars = flexScrollBarVertical
        End If
        
        'Resultset이 가지는 자료를 그리드로 넘긴다.
        .Row = 0
        For i = 0 To .Cols - 1
            .Col = i
            .Text = rs(i).Name
        Next i
        .ColAlignment(1) = flexAlignLeftCenter
        
        While Not rs.EOF
            DoEvents
        
            .Row = .Row + 1
            .RowHeight(.Row) = 290
            For i = 0 To .Cols - 1
                .Col = i
                .Text = CheckNull(rs(i))
            Next i
            If .Rows > 50 Then proProgress.Value = CInt(.Row / (.Rows - 1) * 100)
            lblCount = Format(.Row, "#,###")
            
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        
        If .Rows > 50 Then proProgress.Value = 100
        
        If .FixedRows < .Rows Then
            .HighLight = flexHighlightAlways
        
            '디폴트로 그리드의 첫번째 항목을 선택하여 둔다.
            .Row = 1
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
        End If
        iGridWidth = pnlChoice.Width + iGridWidth + 485
        
        If iGridWidth > m_iLimitForm Then
            Me.Width = iGridWidth
        End If
        
        .Redraw = True
    End With
    proProgress.Visible = False
    
    Screen.MousePointer = vbArrow
End Sub

Private Function Max(Value1, Value2)
    If Value1 > Value2 Then
        Max = Value1
    Else
        Max = Value2
    End If
End Function

Private Sub cmdChoice_Click(Index As Integer)
    Call SetGrid(FL_BY_BTN, CStr(Index))
    
    grdData.SetFocus
End Sub

Private Sub cmdOK_Click()
    If grdData.Rows > 1 Then
        Call SelectData
    Else
        MsgBox LoadResString(111), vbInformation
        cmdSearch.SetFocus
    End If
End Sub

Private Sub cmdSearch_Click()
    If Len(Trim(txtCode)) > 0 Then
        Call SetGrid(FL_BY_CODE, txtCode)
    Else
        Call SetGrid(FL_BY_NAME, txtName)
    End If
    grdData.SetFocus
End Sub

Private Sub Form_Activate()
    If grdData.Rows > 1 Then
        grdData.SetFocus
    Else
        cmdSearch.SetFocus
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        m_bSelected = False
        Me.Visible = False
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmFind = Nothing
End Sub

Private Sub Form_Resize()
    grdData.Move 975, 825, Me.ScaleWidth - pnlChoice.Width - 35, Me.ScaleHeight - pnlFind.Height - 35
    
    proProgress.Width = grdData.Width - 420
End Sub

Private Sub grdData_DblClick()
    Call SelectData
End Sub

Private Sub SelectData()
    Dim i%
    
    On Error Resume Next
    
    If grdData.Rows > 1 Then
        m_bSelected = True
        
        ReDim wData(grdData.Cols - 1)
        With grdData
            For i = 0 To .Cols - 1
                .Col = i
                wData(i) = .Text
            Next i
        End With
        
        Me.Visible = False
    Else
        MsgBox "검색된 내용이 없습니다." & vbCrLf & "검색 후 다시 선택하여 주십시요", vbInformation
        cmdSearch.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call SelectData
    ElseIf KeyAscii = vbKeyEscape Then
        m_bSelected = False
        Me.Visible = False
    End If
End Sub

Private Sub optName_Click(Index As Integer)
    If Index = 0 Then
    
    Else
    
    End If
End Sub

Private Sub txtCode_GotFocus()
    With txtCode
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    cmdSearch.Default = True
End Sub

Private Sub txtCode_LostFocus()
    cmdSearch.Default = False
End Sub

Private Sub txtName_GotFocus()
    With txtName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    cmdSearch.Default = True
End Sub

Private Sub txtName_LostFocus()
    cmdSearch.Default = False
End Sub
