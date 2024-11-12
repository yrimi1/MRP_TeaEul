VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInspectResultByLot 
   ClientHeight    =   9255
   ClientLeft      =   1665
   ClientTop       =   1470
   ClientWidth     =   11865
   Icon            =   "frmInspectResultByLot.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.ComboBox cboExamNo 
      Height          =   300
      Left            =   10095
      Style           =   2  '드롭다운 목록
      TabIndex        =   26
      Top             =   450
      Width           =   870
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7605
      Left            =   0
      TabIndex        =   25
      Top             =   840
      Width           =   11865
      _cx             =   20929
      _cy             =   13414
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
   Begin VB.Frame fraOrder 
      Height          =   795
      Left            =   45
      TabIndex        =   22
      Top             =   -15
      Width           =   1320
      Begin VB.OptionButton optOrder 
         Caption         =   "Order No."
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   210
         Width           =   1170
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "관리 번호"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   495
         Value           =   -1  'True
         Width           =   1110
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   1
      Left            =   7095
      TabIndex        =   15
      Top             =   105
      Width           =   1500
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   300
      Index           =   2
      Left            =   10095
      MaxLength       =   4
      TabIndex        =   14
      Top             =   105
      Width           =   870
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금년"
      Height          =   315
      Index           =   3
      Left            =   2070
      MousePointer    =   99  '사용자 정의
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금일"
      Height          =   315
      Index           =   2
      Left            =   1425
      MousePointer    =   99  '사용자 정의
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "금월"
      Height          =   315
      Index           =   1
      Left            =   2070
      MousePointer    =   99  '사용자 정의
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "전월"
      Height          =   315
      Index           =   0
      Left            =   1425
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색(&F)"
      Height          =   780
      Left            =   11055
      MousePointer    =   99  '사용자 정의
      Style           =   1  '그래픽
      TabIndex        =   0
      ToolTipText     =   "자료 검색"
      Top             =   30
      Width           =   780
   End
   Begin Threed.SSPanel pnlLanguage 
      Height          =   690
      Left            =   15
      TabIndex        =   1
      Top             =   8520
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1217
      _Version        =   196610
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optPrint 
         Caption         =   "영문"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Width           =   690
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "한글"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   3930
      TabIndex        =   8
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   139395073
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   3930
      TabIndex        =   9
      Top             =   450
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   139395073
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   2715
      TabIndex        =   10
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "검사일자"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Index           =   0
      Left            =   8445
      TabIndex        =   12
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   13
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   0
      Top             =   8805
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   3
      Left            =   5880
      TabIndex        =   16
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "관리번호"
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   17
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   4
      Left            =   8820
      TabIndex        =   18
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "LOT No."
         Height          =   240
         Index           =   2
         Left            =   60
         TabIndex        =   19
         Top             =   45
         Width           =   1095
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   8820
      TabIndex        =   27
      Top             =   450
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "검사호기"
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   28
         Top             =   45
         Width           =   1020
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   690
      Left            =   4800
      TabIndex        =   29
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "    엑셀양식인쇄(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "부터"
      Height          =   180
      Index           =   3
      Left            =   5205
      TabIndex        =   21
      Top             =   165
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "까지"
      Height          =   180
      Index           =   2
      Left            =   5205
      TabIndex        =   20
      Top             =   510
      Width           =   360
   End
End
Attribute VB_Name = "frmInspectResultByLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'********************************************************************************************
'변경이력
' 요청 ID : S_201107_태을염직_02
' 요청자 : 김대진 대리
' 요청내용 : 영문 Lot별 검사 결과표 요청
' 변경일자 : 2012.07.12
' 변경내용 : 영문 레포트 추가
'
'********************************************************************************************
Option Explicit

Private Const REPORTFILE_3 = "\Report\InspectResultByLot_K.xls"
Private Const REPORTFILE_4 = "\Report\InspectResultByLot_E.xls"

'S_201107_태을염직_02 에 의한 수정(OLD: InspectResultByLot.rpt)
Private Const REPORTFILE_1 = "\Report\InspectResultByLot_K.rpt"
Private Const REPORTFILE_2 = "\Report\InspectResultByLot_E.rpt"


Private Const BASE_X       As Integer = 150
Private Const BASE_Y       As Integer = 1300
Private Const DEFECT_COUNT As Integer = 50


Private Type TDefect
    Korean  As String
    English As String
    Defect  As String
End Type

Private m_iSortType As Integer
Private m_nSelected As Integer
Dim m_sTotalField(7)  As String             ' 리포트 Title
Dim m_nDefectName(DEFECT_COUNT) As TDefect
Dim m_nPageCnt(1) As Integer



Dim iExcelByPage As Integer                     '한 페이장 출력되는 불량의 ROW수 23개 * 2줄씩
Private Const iDataStartRow As Integer = 9      '엑셀의 맨 첫 페이지의 데이터가 시작되는 행 조일 :9
Private Const EXCEL_ROW As Integer = 63         '엑셀 한 페이지 총 행수(프린트 여백 내) 스카이 :66, 우성:63, FTENE:63, 유창바이오:63

Private nChkRollList As Integer                 'Roll List
Private nChkDate%, sSDate$, sEDate$             '검사일자 체크
Private nChkLotNO%, sLotNo$                     'LOT No
Private nChkExamNO%, sExamNO$                   '검사호기
Private nChkrollID%, SRollID%, ERollID%         'RollNo범위선택
'Private m_iSortType As Integer









Private Sub cmdPrint_Click(Index As Integer)
    Dim i%

    If grdData.Rows = grdData.FixedRows Then Exit Sub


    If m_nSelected = 1 Then
        With grdData
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 1) = flexChecked Then Exit For
            Next i
            PopupMenu PlusMDI.mnuPopup
            Call ReportPrint(PlusMDI.PrintPreview, MakeOrderID(.TextMatrix(i, 3), OM_REDUCE), .TextMatrix(i, 16))
        End With
    ElseIf m_nSelected > 1 Then
        With grdData
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 1) = flexChecked Then
                    .Row = i
                    Call ReportPrint(False, MakeOrderID(.TextMatrix(i, 3), OM_REDUCE), .TextMatrix(i, 16))
                End If
            Next i
        End With
    End If

End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11985, 9660

    Call SetOperate(Me)
    
    cmdPrint(0).Picture = LoadResPicture("PRINT", vbResIcon)
    
    With cboExamNo
        For i = 1 To 10
            .AddItem Format(i, "00") & "호기"
        Next i
        .ListIndex = 0
    End With
    dtpDate(0) = Now
    dtpDate(1) = Now

    Call InitGrid

    Show

    chkSearch(1).Value = vbChecked
    m_nSelected = 0
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Call SetDtpDate(Index, dtpDate(0), dtpDate(1))

    cmdSearch.SetFocus
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If chkSearch(Index) Then
        If Index = 0 Then
            dtpDate(0).SetFocus
        ElseIf Index = 2 Then
            txtSearch(2).SetFocus
        ElseIf Index = 3 Then
            cboExamNo.SetFocus
        End If
    Else
        cmdSearch.SetFocus
    End If
End Sub

Private Sub SSCommand1_Click()

    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    Dim sDate$, eDate$
    
    On Error GoTo ErrHandler
     
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass
    
    If grdData.Rows = grdData.FixedRows Then
        MsgBox "검색 후 인쇄해 주세요", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If chkSearch(0) = vbChecked And chkSearch(1) = vbChecked Then
      If MsgBox("관리번호와 날짜를 함께 검색 시," & vbCrLf _
        & "검사날짜가 다른데이터가 누락되어 출력될수있습니다" & vbCrLf _
        & "이대로 출력하시겠습니까? ", vbQuestion + vbYesNo, "취소 여부 ") = vbNo Then
        Screen.MousePointer = vbDefault
         Exit Sub
      End If
    End If
    
   
    Call ReportPrintExcel(PlusMDI.PrintPreview)

 
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmSubulReport.cmdPrint_Click", Err.Description)

End Sub

Private Sub ReportPrintExcel(bDirect As Boolean)
    
    Dim oInspect As PlusLib2.CInspect
    Dim oCode    As PlusLib2.CCode
    Dim rs       As ADODB.Recordset
    Dim rsQty       As ADODB.Recordset
    Dim rsDefect As ADODB.Recordset
    Dim rsDensity As ADODB.Recordset
    Dim rsCount  As ADODB.Recordset
    Dim rsGradeCnt  As ADODB.Recordset             '등급별 수량
    Dim sParam() As String
    Dim i%, iPoint%, iLoop%
    Dim sReport$, sDate$, sMachineNO$, sLength$, sRoll$, sLot$, sCondition$
 
    Dim sDensity$, nChkRoll%
    Dim nTotalQty#, nBonusQty#, nCutQty#, nRollCnt%
    Dim nSKClss%
    Dim bPoint As Boolean
    Dim nRow%, nCol%, nPage%, nBaseRow%
    Dim sBaseQty$, sLimitQty$, sRejectQty$
    Dim sOrderNO$, sWorkSeq$, sColor$
    Dim nStuffQty As Long, nSampleQty As Long, nCtrlQty As Single, nDemerit As Long, nCalcValue As Single, nRecordCount As Integer
    Dim nGross As Single, Grade As String
    Dim GradeA As Single, GradeB As Single, GradeC As Single, GradeD As Single, GradeE As Single, GradeF As Single
    Dim nWeight As Double           '총 중량
    Dim sOrderID As String
    Dim j3, k3 As Integer
    Dim xlApp   As Excel.Application
    Dim xlBook  As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim oFs         As FileSystemObject
    Dim nDefectQty  As Integer
    Dim lnCheckedQty    As Long
    
    Dim iSQLFieldCnt As Integer                '불량내역이 시작되는 RS열
        Dim iDefectColSpan As Integer          '각 불량의 엑셀에서의 간격
    Dim iDefectColRep As Integer               'SQL의 D,T,E,L 등의 갯수
    Dim iDefectCntByRow As Integer             '한 Row당 불량의 갯수
    
    Dim iExcelDefStartCol As Integer           '엑셀의 불량이 시작되는 열

    Dim iDefCnt, iExcelCol As Integer           '불량갯수 증가
    Dim iInsRow As Integer                      '행 추가 Row

    Dim lnCurRow        As Long
    
    Dim sColorID As Integer

    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
    iExcelDefStartCol = 14
    iSQLFieldCnt = 5                   'SQL의 FieldCount의 값을 가져옴 '26
    iDefectColSpan = 3
    iDefectColRep = 4
    iDefectCntByRow = 15
  
    iExcelByPage = 23 * 2                 '불량의 한페이지당 총 갯수는 23개임(2줄씩)

    Set xlApp = New Excel.Application
    
    If optPrint(0).Value Then
         Set xlBook = xlApp.Workbooks.Open(App.Path & REPORTFILE_3)
    ElseIf optPrint(1).Value Then
         Set xlBook = xlApp.Workbooks.Open(App.Path & REPORTFILE_4)
  
    End If
  
    
    Set oFs = New FileSystemObject
    If Not oFs.FolderExists(App.Path & "\Excel") Then
        oFs.CreateFolder (App.Path & "\Excel")
    End If
    
    
   With grdData

     
     sOrderID = MakeOrderID(.TextMatrix(.Row, .ColIndex("OrderID")), OM_REDUCE)

 
        sReport = App.Path & "\Excel\" & sOrderID & "_Inspect.xls"
        If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
                        
        
        
        '그리드와 출력순 맞추기 위해 값 초기화
        nPage = 1
        nBaseRow = 0
        nRow = 0

            
    For i = .FixedRows To .Rows - 1
    If .Cell(flexcpChecked, i, 1) = flexChecked Then

            
            nStuffQty = 0
            nCtrlQty = 0
            nDemerit = 0
            nCalcValue = 0
            nRecordCount = 0
 
            sOrderNO = .TextMatrix(i, .ColIndex("OrderNo"))                                'Order No
            sOrderID = MakeOrderID(.TextMatrix(i, .ColIndex("OrderID")), OM_REDUCE)        '관리번호
            sColorID = .TextMatrix(i, .ColIndex("OrderSeq"))
            sLotNo = .TextMatrix(i, .ColIndex("LotNo"))
            
            
                 
                
            nChkRollList = 0
            If m_nSelected = 1 Then     '단일 항목 체크일경우
                nChkRollList = 1
            End If

               
                    With xlApp
                    
                        .Worksheets("Form").Activate
                        
                        Set oInspect = New PlusLib2.CInspect
                        oInspect.Connection = g_adoCon
                
                        'xp_Inspect_pResultByColor-핵심데이타 리스트 프로시져
''                        Set rs = oInspect.PrintResultByOrderExcel(sOrderID, nChkDate, sSDate, sEDate, _
''                                            nChkLotNO, sLotNo, nChkExamNO, sExamNO)
                        Set rs = oInspect.PrintResultByLotExcel(sOrderID, sColorID, sLotNo, _
                                              nChkDate, sSDate, sEDate, nChkExamNO, sExamNO)

                        Set oInspect = Nothing

                        nRow = 0
                        nCol = 0
                        
                        '****Excel PageFooter Start========================================================================================
                        Set oInspect = New PlusLib2.CInspect
                        oInspect.Connection = g_adoCon
                        ' 불량내역
                        ' 언어별 불량내역 리스트 가져오기(1:한,2:영어)
                        Set rsDefect = oInspect.GetDefectByLang(IIf(optPrint(0), 1, 2))
                        Set oInspect = Nothing

                        '불량내역 언어별 리스트 FOR문-엑셀파일 하단 추가설명 부분
                        For k3 = 1 To rsDefect.RecordCount
                            If k3 > 49 Then Exit For

                            .Cells(56 + nRow, 2 + nCol) = CStr(rsDefect!Tag) & "-" & CStr(CheckNull(rsDefect!Display))

                            If (k3 Mod 7) = 0 Then  '불량내역은 7개 단위로 행 변환
                                nRow = nRow + 1
                                nCol = 0
                            Else

                                nCol = nCol + 11
                            End If

                            rsDefect.MoveNext
                        Next k3
                        Set rsDefect = Nothing
                        '****Excel PageFooter End========================================================================================

                        '기준치 허용치 검색조건들..
                        '검사일자
                        '검사일자 초기화
                        sCondition = ""
                        If chkSearch(0).Value = vbChecked Then
                            If dtpDate(0) = dtpDate(1) Then
                                If optPrint(0).Value = True Then
                                    sDate = "검사일자 : " & MakeDate(DF_LONG, dtpDate(0))
                                  ElseIf optPrint(1).Value = True Then
                                    sDate = "INSPECT DATE : " & MakeDate(DF_LONG, dtpDate(0))
                               
                                End If
                            Else
                                If optPrint(0).Value = True Then
                                    sDate = "검사일자 : " & MakeDate(DF_LONG, dtpDate(0)) & "~" & MakeDate(DF_LONG, dtpDate(1))
                                ElseIf optPrint(1).Value = True Then
                                    sDate = "INSPECT DATE : " & MakeDate(DF_LONG, dtpDate(0)) & "~" & MakeDate(DF_LONG, dtpDate(1))
                               
                                End If
                            End If
                        Else
                            sDate = ""
                        End If
                        
                        sCondition = sCondition & sDate
                        
                        ' Lot 번호
                        If chkSearch(2).Value = vbChecked Then
                            sLot = IIf(Len(sCondition) > 0, Space(5), "") & "LOT : " & txtSearch(2)
                        Else
                            sLot = ""
                        
                        End If
                        
                        sCondition = sCondition & sLot
                        

                        ' 호기
                        If chkSearch(3).Value = vbChecked Then
                            If optPrint(0).Value Then
                                sMachineNO = IIf(Len(sCondition) > 0, Space(5), "") & "호기 : " & cboExamNo
                            ElseIf optPrint(1).Value Then
                                sMachineNO = IIf(Len(sCondition) > 0, Space(5), "") & "Machine NO : " & Left(cboExamNo, 1)

                            End If
                        Else
                            sMachineNO = ""
                        End If

                        
                        sCondition = sCondition & sMachineNO

    
                        '****Excel PageHeader Start========================================================================================
                        ' markclss
                        .Cells(3, 1) = sCondition
                        
 
          
''                        'S_201203_조일_04 에 의한 수정 - NEW소스
''                        .Cells(3, 73) = Format(Now, "YYYY/MM/DD")                   '발행일자
''                        .Cells(4, 9) = MakeOrderID(sOrderID, OM_EXPAND)             '관리번호
''                        .Cells(4, 30) = grdData.TextMatrix(grdData.Row, IIf(optPrint(0), grdData.ColIndex("Kcustom"), grdData.ColIndex("ECustom")))   '거래처 (4:한글,12:영문)
''                        .Cells(4, 53) = grdData.TextMatrix(grdData.Row, 9)          '수주량
''                        .Cells(4, 71) = IIf(grdData.TextMatrix(grdData.Row, grdData.ColIndex("UnitClss")) = "0", "YARD", "METER")  '단위
''
''                        .Cells(5, 9) = sOrderNO                                    ' ORDER NO
''                        .Cells(5, 30) = grdData.TextMatrix(grdData.Row, 5)         ' 품명
''                        .Cells(5, 53) = grdData.TextMatrix(grdData.Row, 7)         ' 가공명
''                        .Cells(5, 71) = grdData.TextMatrix(grdData.Row, 15)        ' 원단폭
''                        .Cells(6, 9) = Format(grdData.TextMatrix(grdData.Row, 12), "#,##0") & IIf(grdData.TextMatrix(grdData.Row, grdData.ColIndex("UnitClss")) = "0", "Y", "M")      ' 총수량
''                        .Cells(6, 27) = Format(grdData.TextMatrix(grdData.Row, 18), "#,##0.0")  '총 보상
''                        .Cells(6, 44) = Format(grdData.TextMatrix(grdData.Row, 11), "#,##0")    '총 절수
''                        .Cells(6, 62) = Format(grdData.TextMatrix(grdData.Row, 19), "#,##0")    '총 견본
''                        .Cells(6, 77) = Format(grdData.TextMatrix(grdData.Row, 20), "#,##0")    '난단
                        
                        
                        '2024.11.08 체크한 목록 모두한번에 출력을 위해 수정함
                        .Cells(3, 73) = Format(Now, "YYYY/MM/DD")                   '발행일자
                        .Cells(4, 9) = MakeOrderID(sOrderID, OM_EXPAND)              '관리번호
                        .Cells(4, 30) = grdData.TextMatrix(i, IIf(optPrint(0), grdData.ColIndex("Kcustom"), grdData.ColIndex("ECustom")))   '거래처 (4:한글,12:영문)
                        .Cells(4, 53) = grdData.TextMatrix(i, 9)          '수주량
                        .Cells(4, 71) = IIf(grdData.TextMatrix(i, grdData.ColIndex("UnitClss")) = "0", "YARD", "METER")  '단위
                       
                        .Cells(5, 9) = sOrderNO                                    ' ORDER NO
                        .Cells(5, 30) = grdData.TextMatrix(i, 5)         ' 품명
                        .Cells(5, 53) = grdData.TextMatrix(i, 7)         ' 가공명
                        .Cells(5, 71) = grdData.TextMatrix(i, 15)        ' 원단폭
                        .Cells(6, 9) = Format(grdData.TextMatrix(i, 12), "#,##0") & IIf(grdData.TextMatrix(i, grdData.ColIndex("UnitClss")) = "0", "Y", "M")      ' 총수량
                        .Cells(6, 27) = Format(grdData.TextMatrix(i, 18), "#,##0.0")  '총 보상
                        .Cells(6, 44) = Format(grdData.TextMatrix(i, 11), "#,##0")    '총 절수
                        .Cells(6, 62) = Format(grdData.TextMatrix(i, 19), "#,##0")    '총 견본
                        .Cells(6, 77) = Format(grdData.TextMatrix(i, 20), "#,##0")    '난단
                        
                        
                        
                        ' 등급별 수량
                        Set oInspect = New PlusLib2.CInspect
                        oInspect.Connection = g_adoCon
'
''                        Set rsGradeCnt = oInspect.GetGradeQtyByOrder(sOrderID, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
''                            IIf(chkSearch(2), 1, 0), txtSearch(2), IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
                            
                            
                        Set rsGradeCnt = oInspect.GetGradeQtyByLot(sOrderID, sColorID, sLotNo, _
                                              nChkDate, sSDate, sEDate, nChkExamNO, sExamNO)
                            

                        '합격/불합격 수량-각 등급별 수량 A-D밖에 없음
                        For iLoop = 0 To rsGradeCnt.RecordCount - 1

                            If rsGradeCnt!GradeID = "1" Then               '합격-A등륵
                                .Cells(7, 11) = IIf(IsNull(rsGradeCnt!GradeCount), "0", Format(rsGradeCnt!GradeCount, "#,##0"))
                            ElseIf rsGradeCnt!GradeID = "2" Then           '불량 B등급
                                .Cells(7, 24) = IIf(IsNull(rsGradeCnt!GradeCount), "0", Format(rsGradeCnt!GradeCount, "#,##0"))
                            ElseIf rsGradeCnt!GradeID = "3" Then           '불량 C등급
                                .Cells(7, 36) = IIf(IsNull(rsGradeCnt!GradeCount), "0", Format(rsGradeCnt!GradeCount, "#,##0"))
                            ElseIf rsGradeCnt!GradeID = "4" Then           '불량 F등급
                                .Cells(7, 73) = IIf(IsNull(rsGradeCnt!GradeCount), "0", Format(rsGradeCnt!GradeCount, "#,##0"))           'F등급위치
                            End If

                            rsGradeCnt.MoveNext
                        Next iLoop
                        rsGradeCnt.Close
                        Set rsGradeCnt = Nothing
                        
                        '****Excel PageHeader End========================================================================================
                        
                        nBaseRow = GetExcelBaseRow(nPage)
                        nRow = 0
                        nRecordCount = rs.RecordCount
                        
                        '페이지 추가
                        Call InsertExcelForm(xlApp, nPage)
                        .Worksheets("Report").Activate
                        For j3 = 0 To rs.RecordCount - 1
                        
                             If nRow > iExcelByPage - 2 Then            '46줄 이상일 경우 페이지 증가(OLD : nRow > 46)
                                nPage = nPage + 1
                                Call InsertExcelForm(xlApp, nPage)
                                nBaseRow = GetExcelBaseRow(nPage)
                                nRow = 0
                             End If


 
   
                            lnCurRow = iDataStartRow + nBaseRow + nRow
                            '색상일 경우
                            If CheckNull(rs!Cls) = 1 Then
                                     .Range("A" & lnCurRow & ":I" & lnCurRow + 1).Select
                                With .Selection
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlTop
                                    .WrapText = True
                                    .Orientation = 0
                                    .AddIndent = False
                                    .ShrinkToFit = False
                                    .Borders(xlEdgeRight).LineStyle = xlNone
                                    .Font.Size = 10
                                    .WrapText = False
                                    .ShrinkToFit = True
                                    .Font.Bold = True
                                End With
                                .Selection.Merge

                                    .Range("J" & lnCurRow & ":BG" & lnCurRow + 1).Select
                                With .Selection
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlTop
                                    .WrapText = True
                                    .Orientation = 0
                                    .AddIndent = False
                                    .ShrinkToFit = False
                                    .Borders(xlEdgeRight).LineStyle = xlNone
                                    .Font.Size = 10
                                    .WrapText = False
                                    .ShrinkToFit = True
                                    .Font.Bold = True
                                End With
                                .Selection.Merge
                                
                                
                                .Cells(lnCurRow, 1) = "Lot : " & CheckNull(rs!LotNo)             'Lotno
                                .Cells(lnCurRow, 10) = CheckNull(rs!Color)                       'Color



                            Else
                                .Cells(lnCurRow, 1) = CheckNull(rs!LotNo)              'LOT No
                                .Cells(lnCurRow, 4) = CheckNull(rs!RollNo)             '절번호
                                '.Cells(iDataStartRow + nBaseRow + nRow, 7) = rs!ExamNO          '검사호기
                                .Cells(lnCurRow, 7) = CheckNull(rs!Person)              '검사자 대신 ExamNo
                                .Cells(lnCurRow, 10) = CheckNull(rs!StuffQty)                                '투입량

                                .Cells(lnCurRow, 59) = CheckNull(Format(rs!SampleQty, "0"))                 '견본
                                
                                If Trim(CheckNull(rs!GradeID)) = "" Then        '등급
                                    .Cells(lnCurRow, 73) = ""
                                Else
                                    If rs!GradeID <> 4 Then
                                        .Cells(lnCurRow, 73) = Chr(Asc("@") + CheckNull(CInt(rs!GradeID)))          'ABC(1~3) 등급
                                    Else
                                        .Cells(lnCurRow, 73) = "F"                                       'F등급(4)
                                    End If
                                End If
    
                                '대표불량
                                .Cells(lnCurRow, 76) = IIf(optPrint(0), CheckNull(rs!DefectKor), CheckNull(rs!DefectEng))
                                
                                
    
                                iExcelCol = 0
                                iInsRow = 0
    
    
                                nDefectQty = CheckNum(rs!DefectQty)         '총 불량 수량
                                iExcelCol = 0
    
    
    
                                For iDefCnt = 1 To nDefectQty        'D,E,T,M (4개 단위로 증가함)
    
                                    If nRow > iExcelByPage - 2 Then            '44(iExcelByPage - 2) 줄 페이지 증가(OLD : nRow > 46)
                                       nPage = nPage + 1
                                       Call InsertExcelForm(xlApp, nPage)
                                       nBaseRow = GetExcelBaseRow(nPage)
                                       nRow = 0
                                    End If
    
    ''                                '//불량 내역
                                    'S_201412_조일_04 에 의한 수정( 9 대신 iDataStartRow, 10대신 iDataStartRow+1 사용)
                                    '항목 시작열이 iExcelDefStartCol(14)임 - 4개가 기본셋트 이므로 X 4(iDefectColRep)
                            
                                    .Cells(lnCurRow, iExcelDefStartCol + iExcelCol) = Format(CheckNull(rs(iSQLFieldCnt + (iDefCnt))), "#,##0")    '불량위치
    
    ''                              '불량아랫줄
                                    'S_201901_태을염직_02 에 의한 수정 : 30 -> 60
                                    .Cells(lnCurRow + 1, iExcelDefStartCol + iExcelCol) = CheckNull(rs(iSQLFieldCnt + 60 + (iDefCnt)))      '불량태그
 
    
                                    If (iExcelDefStartCol + iExcelCol) >= iExcelDefStartCol + (iDefectCntByRow - 1) * iDefectColSpan Then         '1 줄 인쇄 완료되면 다음줄로 계속 인쇄(기본 불량 한줄에 15개) (14열부터 56열)
                                        iExcelCol = 0
    
                                        If nDefectQty Mod 15 <> 0 Or iDefCnt <> nDefectQty Then                      '15의 배수로 끝날경우는 다음행 증가 시키지 않음
                                            nRow = nRow + 2     '행의 증가는 이렇게만 지정
                                        End If
    ''                                    iTimes = iTimes - 1       '//불량이 15개의 곱인경우(15,30,45,60 등의 경우 한줄이 더 생기는 오류 체크위함)
                                    Else
    '                                    iExcelCol = iExcelCol + 3
                                        iExcelCol = iExcelCol + iDefectColSpan
                                    End If
    
                                Next iDefCnt
                            End If
                               
                               
                               
'''''                            .Cells(lnCurRow, 10) = CheckNull(rs!StuffQty)                                '투입량
'''''
'''''                            .Cells(lnCurRow, 59) = CheckNull(Format(rs!SampleQty, "0"))                 '견본
                            .Cells(lnCurRow, 62) = CheckNull(Format(rs!CutQty, "0"))                     '난단
                            .Cells(lnCurRow, 64) = CheckNull(Format(rs!CtrlQty, "#,##0"))                 '실수량

 
                            .Cells(lnCurRow, 68) = CheckNull(Format(rs!LossQty, "0"))                    '보상
                            .Cells(lnCurRow, 70) = CheckNull(rs!DefectQty)                               '불량갯수
 

'
                            nStuffQty = nStuffQty + rs!StuffQty                 '총 투입량
                            nWeight = nWeight + rs!CtrlQty                      '총 중량



                            nRow = nRow + 2

                            rs.MoveNext
                        Next j3
                        
                        rs.Close

                        'nRow 44 초과일경우 페이지 추가
                        If nRow > iExcelByPage - 2 Then            '44(iExcelByPage - 2) 줄 페이지 증가(OLD : nRow > 46)
                            nPage = nPage + 1
                            Call InsertExcelForm(xlApp, nPage)
                            nBaseRow = GetExcelBaseRow(nPage)
                            nRow = 0
                        End If
                        Call ExcelTotal(xlApp, nPage, nBaseRow, nRow)  '////음영처리////
                                
                        If optPrint(0).Value Then
                            .Cells(iDataStartRow + nBaseRow + nRow, 1) = "합계 : "
                        ElseIf optPrint(1).Value Then
                            .Cells(iDataStartRow + nBaseRow + nRow, 1) = "TOTAL : "

                        End If
                        .Cells(iDataStartRow + nBaseRow + nRow, 7) = Format(nStuffQty, "#,###")         '총 투입량
                        

                    End With            '// With xlApp 의 끝
                
                nPage = nPage + 1   '기존 소스에 변경-칼라 바뀔때 페이지 추가
        
            End If
        Next i

    End With        '//With grdData의 끝
    
    
    Set oFs = Nothing

    Call xlBook.SaveAs(sReport)

    '//미리보기 출력체크**********
    If Not bDirect Then         '미리보기 없음
        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1

        Call ProcessClose("XLMAIN")
    Else                        '미리보기 있음
        xlApp.WindowState = xlMaximized
        xlApp.Application.Visible = True
        xlApp.ActiveWindow.SelectedSheets.PrintPreview
    End If
    Set rs = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set oFs = Nothing
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:

    Screen.MousePointer = vbDefault
    Set rs = Nothing
    
    Set rsDefect = Nothing
    Set rsDensity = Nothing
    Set rsCount = Nothing
    
    Set oCode = Nothing
    Set oInspect = Nothing

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set oFs = Nothing

    Call ProcessClose("XLMAIN")
    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub cmdSearch_Click()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As Recordset
    Dim i%, iNowRow%

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetResultByLot(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(1), IIf(optOrder(0), 2, 1), 0), IIf(optOrder(0), txtSearch(1), MakeOrderID(txtSearch(1), OM_REDUCE)), _
        IIf(chkSearch(2), 1, 0), txtSearch(2), IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    Set oInspect = Nothing

    With grdData
        .Redraw = flexRDNone

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & vbTab & rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!kCustom & vbTab & rs!Article & vbTab & MakeDate(DF_LONG, rs!DvlyDate) & vbTab & _
                rs!WorkName & vbTab & rs!Color & vbTab & CheckNum(rs!ColorQty) & vbTab & _
                rs!LotNo & vbTab & CheckNum(rs!PassRoll) & vbTab & CheckNum(rs!PassQty) & vbTab & CheckNull(rs!ECustom) & vbTab & _
                rs!UnitClss & vbTab & rs!WorkWidth & vbTab & rs!OrderSeq & vbTab & CheckNull(rs!DesignNO) & vbTab & _
                CheckNum(rs!LossQty) & vbTab & CheckNum(rs!SampleQty) & vbTab & CheckNum(rs!CutQty)

            If (i Mod 2) = 0 Then
                .Row = .FixedRows + i - 1
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = COLOR_GRIDROW
            End If

            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > iNowRow, iNowRow, .Rows - 1)
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If

        m_nSelected = 0

        .Redraw = flexRDDirect

        .SetFocus
    End With

    Screen.MousePointer = vbArrow

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub grdData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdData
        If .Rows = .FixedRows Or .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub

        If .Cell(flexcpChecked, .Row, 1) = flexUnchecked Then
            .Cell(flexcpChecked, .Row, 1) = flexChecked
            m_nSelected = m_nSelected + 1
        Else
            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
            m_nSelected = m_nSelected - 1
        End If
    End With
End Sub

Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        chkSearch(1).Caption = "Order No"
    Else
        chkSearch(1).Caption = "관리번호"
    End If

    cmdSearch.SetFocus
End Sub

Private Sub optPrint_Click(Index As Integer)
    cmdPrint(0).SetFocus
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    With grdData
        .Cols = 21
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .TextArray(1) = "선택":         .ColWidth(1) = 300:             .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "Order No.":    .ColWidth(2) = 1350:            .ColAlignment(2) = flexAlignLeftCenter:     .ColKey(2) = "OrderNo"
        .TextArray(3) = "관리번호":     .ColWidth(3) = 1225:            .ColAlignment(3) = flexAlignCenterCenter:   .ColKey(3) = "OrderID"
        .TextArray(4) = "거래처":       .ColWidth(4) = 1170:            .ColAlignment(4) = flexAlignLeftCenter:     .ColKey(4) = "Kcustom"
        .TextArray(5) = "품명":         .ColWidth(5) = 1800:            .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "납기일자":     .ColWidth(6) = 990:             .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "가공구분":     .ColWidth(7) = 450:             .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "색상명":       .ColWidth(8) = 1530:            .ColAlignment(8) = flexAlignLeftCenter
        .TextArray(9) = "수주수량":     .ColWidth(9) = 830:             .ColAlignment(9) = flexAlignRightCenter:                                .ColFormat(9) = "#,###"
        .TextArray(10) = "LotNo":       .ColWidth(10) = 450:            .ColAlignment(10) = flexAlignRightCenter:   .ColKey(10) = "LotNo"
        .TextArray(11) = "검사절수":    .ColWidth(11) = 450:            .ColAlignment(11) = flexAlignRightCenter:                               .ColFormat(10) = "#,###"
        .TextArray(12) = "검사수량":    .ColWidth(12) = 830:            .ColAlignment(12) = flexAlignRightCenter:                               .ColFormat(11) = "#,###"
        
        .TextArray(13) = "거래처(영)":  .ColWidth(13) = 0:                                                          .ColKey(13) = "ECustom"
        .TextArray(14) = "수량단위":    .ColWidth(14) = 0:                                                          .ColKey(14) = "UnitClss"
        
        .TextArray(15) = "생지폭":      .ColWidth(15) = 0
        .TextArray(16) = "색상순위":    .ColWidth(16) = 0:                                                         .ColKey(16) = "OrderSeq"
        .TextArray(17) = "DesignNo":    .ColWidth(17) = 0
        .TextArray(18) = "보상수량":    .ColWidth(18) = 0:                                                                                      .ColFormat(17) = "#.#"
        .TextArray(19) = "견본수량":    .ColWidth(19) = 0
        .TextArray(20) = "난단수량":    .ColWidth(20) = 0

        .ColDataType(1) = flexDTBoolean

        .Redraw = flexRDDirect
    End With
End Sub

Private Sub ReportPrint(bDirect As Boolean, sOrderID As String, nOrderSeq As Integer)
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As ADODB.Recordset
    Dim sParam() As String
    Dim i%, iPoint%, iLoop%

    On Error GoTo ErrHandler

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    ' footer
    Set rs = oInspect.GetDefectByLang(IIf(optPrint(0), 1, 2))
    Set oInspect = Nothing

    ReDim sParam(44)

    For i = 0 To rs.RecordCount - 1
        sParam(i) = CStr(rs!Tag) & "-" & CStr(CheckNull(rs!Display))

        rs.MoveNext
    Next i
    rs.Close
    Set rs = Nothing

    Do While i <= 44
        sParam(i) = " "
        i = i + 1
    Loop

    ReDim Preserve sParam(58)

    ' MarkClss
    If optPrint(0) Then
        If dtpDate(0) = dtpDate(1) Then
            sParam(i) = "검사일자 : " & Format(dtpDate(0), "YYYY년 MM월 DD일") & Space(5) & _
                IIf(chkSearch(3), "호기 : " & cboExamNo, "전체 호기")
        Else
            sParam(i) = "검사일자 : " & Format(dtpDate(0), "YYYY년 MM월 DD일") & " ~ " & Format(dtpDate(1), "YYYY년 MM월 DD일") & Space(5) & _
                IIf(chkSearch(3), "호기 : " & cboExamNo, "전체 호기")
        End If
    Else
        If dtpDate(0) = dtpDate(1) Then
            sParam(i) = "INSPECTION DATE : " & Format(dtpDate(0), "YYYY/MM/DD") & Space(5) & _
                IIf(chkSearch(3), "Machine No : " & Left(cboExamNo, 2) & "Mc", "Total Machine")
        Else
            sParam(i) = "INSPECTION DATE : " & Format(dtpDate(0), "YYYY/MM/DD") & " ~ " & Format(dtpDate(1), "YYYY/MM/DD") & Space(5) & _
                IIf(chkSearch(3), "Machine No : " & Left(cboExamNo, 2) & "Mc", "Total Machine")
        End If
    End If

    i = i + 1

    ' Header
    iPoint = grdData.Cols * grdData.Row
    sParam(i) = MakeOrderID(sOrderID, OM_EXPAND)
    sParam(i + 1) = grdData.TextArray(iPoint + IIf(optPrint(0), 4, 13))
    sParam(i + 2) = grdData.TextArray(iPoint + 8)
    sParam(i + 3) = IIf(grdData.TextArray(iPoint + 13) = "0", "YARD", "METER")
    sParam(i + 4) = grdData.TextArray(iPoint + 2)
    sParam(i + 5) = grdData.TextArray(iPoint + 5)
    sParam(i + 6) = grdData.TextArray(iPoint + 7)
    sParam(i + 7) = grdData.TextArray(iPoint + 15)
    sParam(i + 8) = Format(grdData.TextArray(iPoint + 12), "#,##0") & IIf(grdData.TextArray(iPoint + 14) = "0", "Y", "M")
    sParam(i + 9) = Format(grdData.TextArray(iPoint + 18), "#,##0")
    sParam(i + 10) = grdData.TextArray(iPoint + 11)
    sParam(i + 11) = Format(grdData.TextArray(iPoint + 19), "#,##0")
    sParam(i + 12) = Format(grdData.TextArray(iPoint + 20), "#,##0")
    
    i = i + 13
    
    ReDim Preserve sParam(66)
    ' title
    If optPrint(0) Then
        sParam(i) = "검사 결과표"
    Else
        sParam(i) = "INSPECTION REPORT"
    End If
    ' companyname
    sParam(i + 1) = CompanyName

    i = i + 2

    ' GradeCount
    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetGradeQtyByColor(sOrderID, nOrderSeq, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), _
        MakeDate(DF_SHORT, dtpDate(1)), 1, grdData.TextArray(iPoint + 10), _
        IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))

    For iLoop = 0 To rs.RecordCount - 1
        sParam(i + iLoop) = IIf(IsNull(rs!GradeCount), "0", rs!GradeCount)
        
        rs.MoveNext
    Next iLoop
    rs.Close
    Set rs = Nothing

    Set rs = oInspect.PrintResultByColor(sOrderID, nOrderSeq, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), _
        MakeDate(DF_SHORT, dtpDate(1)), 1, grdData.TextArray(iPoint + 10), _
        IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    Set oInspect = Nothing

    Call PrintReport(IIf(optPrint(0), REPORTFILE_1, REPORTFILE_2), rs, sParam, bDirect)

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

'S_201211_태을염직_03 에 의한 추가
Private Function InsertExcelForm(xlApp As Excel.Application, nPage As Integer)
 Dim i%, nBaseRow%

    nBaseRow = GetExcelBaseRow(nPage)
    With xlApp
        .Sheets("Form").Select

        .Rows("1:" & CStr(EXCEL_ROW)).Select
        .Selection.Copy

        .Sheets("Report").Select
        .Rows(CStr(nBaseRow + 1) & ":" & CStr(nBaseRow + 1)).Select
        .Selection.Insert Shift:=xlDown
    End With
End Function

Private Function GetExcelBaseRow(nPage)
    GetExcelBaseRow = (nPage - 1) * EXCEL_ROW
End Function

Private Function ExcelTotal(xlApp As Excel.Application, nPage As Integer, nBaseRow As Integer, nRow As Integer)

    On Error GoTo Err_Rtn
    
    With xlApp
        'S_201302_조일_01 에 의한 수정-OLD소스
        '.Range("A" & 9 + nBaseRow + nRow & ":F" & 9 + nBaseRow + 47).Select
        
        'S_201302_조일_01 에 의한 수정-NEW소스
        'iDataStartRow + nBaseRow + nRow => Merge 할 From 행
        'iDataStartRow + nBaseRow + iExcelByPage - 1   => Merge 할 To 행
''        .Range(GF_Excel_CA(1) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(6) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '슷자를 Excel Column 영문자로 변경 Range 설정
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 1), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 6)).Select         'R1C1 스타일 참조 주소
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Font.Size = 10
            .WrapText = False
            .ShrinkToFit = True
            .Font.Bold = True
        End With
        .Selection.Merge

        'S_201302_조일_01 에 의한 수정-OLD소스
        '.Range("G" & 9 + nBaseRow + nRow & ":M" & 9 + nBaseRow + 47).Select
        'S_201302_조일_01 에 의한 수정-NEW소스
''        .Range(GF_Excel_CA(7) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(13) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '슷자를 Excel Column 영문자로 변경 Range 설정
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 7), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 13)).Select         'R1C1 스타일 참조 주소
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Font.Size = 10
            .WrapText = False
            .ShrinkToFit = True
            .Font.Bold = True
        End With
        .Selection.Merge

        'S_201302_조일_01 에 의한 수정-OLD소스
        '.Range("N" & 9 + nBaseRow + nRow & ":BF" & 9 + nBaseRow + 47).Select
        'S_201302_조일_01 에 의한 수정-NEW소스
''        .Range(GF_Excel_CA(14) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(58) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select '슷자를 Excel Column 영문자로 변경 Range 설정
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 14), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 58)).Select         'R1C1 스타일 참조 주소
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Font.Size = 10
            .WrapText = False
            .ShrinkToFit = True
            .Font.Bold = True
        End With
        .Selection.Merge
        .Selection.Interior.ColorIndex = xlNone
        
         
        'S_201302_조일_01 에 의한 수정-OLD소스
        '.Range("BG" & 9 + nBaseRow + nRow & ":BK" & 9 + nBaseRow + 47).Select
        'S_201302_조일_01 에 의한 수정-NEW소스
''        .Range(GF_Excel_CA(59) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(63) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '슷자를 Excel Column 영문자로 변경 Range 설정
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 59), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 63)).Select         'R1C1 스타일 참조 주소
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Font.Size = 10
            .WrapText = False
            .ShrinkToFit = True
            .Font.Bold = True
        End With
        .Selection.Merge

        
        'S_201302_조일_01 에 의한 수정-OLD소스
        '.Range("BL" & 9 + nBaseRow + nRow & ":BO" & 9 + nBaseRow + 47).Select
        'S_201302_조일_01 에 의한 수정-NEW소스
''        .Range(GF_Excel_CA(64) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(67) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '슷자를 Excel Column 영문자로 변경 Range 설정
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 64), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 67)).Select         'R1C1 스타일 참조 주소
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Font.Size = 10
            .WrapText = False
            .ShrinkToFit = True
            .Font.Bold = True
        End With
        .Selection.Merge
               
        'S_201302_조일_01 에 의한 수정-OLD소스
        '.Range("BP" & 9 + nBaseRow + nRow & ":BT" & 9 + nBaseRow + 47).Select
        'S_201302_조일_01 에 의한 수정-NEW소스
''        .Range(GF_Excel_CA(68) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(72) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '슷자를 Excel Column 영문자로 변경 Range 설정
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 68), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 72)).Select         'R1C1 스타일 참조 주소
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Font.Size = 10
            .WrapText = False
            .ShrinkToFit = True
            .Font.Bold = True
        End With
        .Selection.Merge
        
        
        'S_201302_조일_01 에 의한 수정-OLD소스
        '.Range("BU" & 9 + nBaseRow + nRow & ":CF" & 9 + nBaseRow + 47).Select
        'S_201302_조일_01 에 의한 수정-NEW소스
''        .Range(GF_Excel_CA(73) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(84) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '슷자를 Excel Column 영문자로 변경 Range 설정
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 73), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 84)).Select         'R1C1 스타일 참조 주소
        With .Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Font.Size = 10
            .WrapText = False
            .ShrinkToFit = True
            .Font.Bold = True
        End With
        .Selection.Merge

    
    End With
    
    Exit Function
    
Err_Rtn:
    If Err.Number <> 0 Then MsgBox Err.Number & "," & Err.Description, vbCritical, "[ExcelTotal]"
End Function


