VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInspectResultByOrder 
   ClientHeight    =   9255
   ClientLeft      =   4080
   ClientTop       =   2685
   ClientWidth     =   11865
   Icon            =   "frmInspectResultByOrder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   Begin VB.ComboBox cboExamNo 
      Height          =   300
      Left            =   10035
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   27
      Top             =   450
      Width           =   930
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7605
      Left            =   30
      TabIndex        =   26
      Top             =   870
      Width           =   11865
      _cx             =   20929
      _cy             =   13414
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      TabIndex        =   23
      Top             =   -15
      Width           =   1320
      Begin VB.OptionButton optOrder 
         Caption         =   "���� ��ȣ"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   480
         Value           =   -1  'True
         Width           =   1110
      End
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
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   1  '������ ����
      Height          =   300
      Index           =   2
      Left            =   10020
      MaxLength       =   4
      TabIndex        =   16
      Top             =   105
      Width           =   930
   End
   Begin VB.TextBox txtSearch 
      Height          =   300
      Index           =   1
      Left            =   7095
      TabIndex        =   15
      Top             =   105
      Width           =   1500
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "�ݳ�"
      Height          =   315
      Index           =   3
      Left            =   2070
      MousePointer    =   99  '����� ����
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "����"
      Height          =   315
      Index           =   2
      Left            =   1425
      MousePointer    =   99  '����� ����
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "�ݿ�"
      Height          =   315
      Index           =   1
      Left            =   2070
      MousePointer    =   99  '����� ����
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin VB.CommandButton cmdTerm 
      Caption         =   "����"
      Height          =   315
      Index           =   0
      Left            =   1425
      MousePointer    =   99  '����� ����
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   105
      Width           =   615
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "�˻�(&F)"
      Height          =   780
      Left            =   11055
      MousePointer    =   99  '����� ����
      Style           =   1  '�׷���
      TabIndex        =   1
      ToolTipText     =   "�ڷ� �˻�"
      Top             =   30
      Width           =   780
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   6720
      TabIndex        =   0
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      ����(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlLanguage 
      Height          =   690
      Left            =   45
      TabIndex        =   2
      Top             =   8520
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1217
      _Version        =   196610
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optPrint 
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   420
         Width           =   690
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "�ѱ�"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   3
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
      TabIndex        =   9
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   139919361
      CurrentDate     =   36871
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   1
      Left            =   3930
      TabIndex        =   10
      Top             =   450
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      Format          =   139919361
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   2
      Left            =   2715
      TabIndex        =   11
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "�˻�����"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8445
      TabIndex        =   13
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      �μ�(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10170
      TabIndex        =   14
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      �ݱ�(&X)"
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
      Index           =   6
      Left            =   5880
      TabIndex        =   17
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "������ȣ"
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   18
         Top             =   45
         Width           =   1080
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   7
      Left            =   8760
      TabIndex        =   19
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
         TabIndex        =   20
         Top             =   45
         Width           =   1095
      End
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   8760
      TabIndex        =   28
      Top             =   450
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "�˻�ȣ��"
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   29
         Top             =   45
         Width           =   1020
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   690
      Left            =   4800
      TabIndex        =   30
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "    ��������μ�(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   5
      Left            =   5205
      TabIndex        =   22
      Top             =   510
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   4
      Left            =   5205
      TabIndex        =   21
      Top             =   165
      Width           =   360
   End
End
Attribute VB_Name = "frmInspectResultByOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'20151023 ������ ���� �μ��� �߰�.
'2019.01.11, ������, S_201901_��������_02, ���ֺ� ��ȸ - ��������μ�� �� �̻��ϰ� �����°� ����
'2019.01.16, ������, S_201901_��������_04, ���ֺ� ��ȸ - ���ֺ� ��ȸ - ��������μ�(����) �����ߴ°� ����

Option Explicit

Private Const REPORTFILE_3 = "\Report\InspectResultByOrder_K.xls"
Private Const REPORTFILE_4 = "\Report\InspectResultByOrder_E.xls"

Private Const REPORTFILE_1 = "\Report\InspectResultByOrder_K.rpt"
Private Const REPORTFILE_2 = "\Report\InspectResultByOrder_E.rpt"

 
Private Const BASE_X       As Integer = 150
Private Const BASE_Y       As Integer = 1300
Private Const DEFECT_COUNT As Integer = 50


Private Type TDefect
    Korean  As String
    English As String
    Defect  As String
End Type
Dim iExcelByPage As Integer                 '�� ������ ��µǴ� �ҷ��� ROW�� 23�� * 2�پ�
Private Const iDataStartRow As Integer = 9             '������ �� ù �������� �����Ͱ� ���۵Ǵ� �� ���� :9
Private Const EXCEL_ROW As Integer = 63             '���� �� ������ �� ���(����Ʈ ���� ��) ��ī�� :66, �켺:63, FTENE:63, ��â���̿�:63
   
Private m_iSortType As Integer
Private m_nSelected As Integer
Dim m_sTotalField(7)  As String             ' ����Ʈ Title
Dim m_nDefectName(DEFECT_COUNT) As TDefect
Dim m_nPageCnt(1) As Integer

Private nChkRollList As Integer                   'Roll List
Private nChkDate%, sSDate$, sEDate$                 '�˻����� üũ
Private nChkLotNO%, sLotNo$                         'LOT No
Private nChkExamNO%, sExamNO$                       '�˻�ȣ��
Private nChkrollID%, SRollID%, ERollID%             'RollNo��������
'Private m_iSortType As Integer

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11985, 9660

    Call SetOperate(Me)
    
    With cboExamNo
        For i = 1 To 10
            .AddItem Format(i, "00") & "ȣ��"
        Next i
        .ListIndex = 0
    End With
    
    dtpDate(0) = Now
    dtpDate(1) = Now
    cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)

    Call InitGrid

    Show
    
    chkSearch(0).Value = vbChecked
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
    
   ' If grdData.Rows = grdData.FixedRows Then Exit Sub
    If grdData.Rows = grdData.FixedRows Then
        MsgBox "�˻� �� �μ��� �ּ���", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
   
    Call ReportPrintExcel(PlusMDI.PrintPreview)
    'Call MakeExcelInspectResultByColor
 
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
    Dim rsGradeCnt  As ADODB.Recordset             '��޺� ����
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
    Dim nWeight As Double           '�� �߷�
     Dim sOrderID As String
''    Dim nChkRollList As Integer
    Dim j3, k3 As Integer
    Dim xlApp   As Excel.Application
    Dim xlBook  As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim oFs         As FileSystemObject
    Dim nDefectQty  As Integer
    Dim lnCheckedQty    As Long
    
    Dim iSQLFieldCnt As Integer              '�ҷ������� ���۵Ǵ� RS��(��ī�� : 32,�켺�����:25,FTENE : 26, ��â���̿� : 26)
        Dim iDefectColSpan As Integer            '�� �ҷ��� ���������� ���� (��ī��:3,�켺�����:3,FTENE:3, ��â���̿� : 3)
    Dim iDefectColRep As Integer             'SQL�� D,T,E,L ���� ����(��ī��4, �켺�����:4,FTENE:4, ��â���̿� : 4)
    Dim iDefectCntByRow As Integer             '�� Row�� �ҷ��� ����(��ī��:14,�켺�����:15,FTENE:15, ��â���̿� : 15)
    
    Dim iExcelDefStartCol As Integer            '������ �ҷ��� ���۵Ǵ� ��
''    Dim iExcelByPage As Integer                 '�� ������ ��µǴ� �ҷ��� ROW�� 23�� * 2�پ�
    
    Dim iDefCnt, iExcelCol As Integer           '�ҷ����� ����
    Dim iInsRow As Integer                      '�� �߰� Row
''    Dim iTimes As Integer                       '�ҷ��� 15���� ���ΰ��(15,30,45,60 ���� ��� ������ �� ����� ���� üũ����)
    Dim lnCurRow        As Long
    On Error GoTo ErrHandler
    
 Screen.MousePointer = vbHourglass
    
    iExcelDefStartCol = 14
    iSQLFieldCnt = 5                 'SQL�� FieldCount�� ���� ������   'S_201901_��������_02 �� ���� ���� : 35 -> 5
    iDefectColSpan = 3
    iDefectColRep = 4
    iDefectCntByRow = 15
  
    iExcelByPage = 23 * 2                 '�ҷ��� ���������� �� ������ 23����(2�پ�)

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
        'Order No  'S_201203_����_04 ����
   '  sOrderID = MakeOrderID(.TextMatrix(.Row, .ColIndex("OrderID"), OM_REDUCE))
     sOrderID = MakeOrderID(.TextMatrix(.Row, 2), OM_REDUCE)
     
 
        sReport = App.Path & "\Excel\" & sOrderID & "_Inspect.xls"
        If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
                        
        '�׸���� ��¼� ���߱� ���� �� �ʱ�ȭ
        nPage = 1
        nBaseRow = 0
        nRow = 0

            nStuffQty = 0
            nCtrlQty = 0
            nDemerit = 0
            nCalcValue = 0
            nRecordCount = 0
 
            sOrderNO = .TextMatrix(.Row, 1)                                'Order No
            sOrderID = MakeOrderID(.TextMatrix(.Row, .ColIndex("OrderID")), OM_REDUCE)        '������ȣ
                 
                
            nChkRollList = 0
            
            nChkRollList = 1
               
                    With xlApp
                    
                        .Worksheets("Form").Activate
                        
                        Set oInspect = New PlusLib2.CInspect
                        oInspect.Connection = g_adoCon
                
                        'xp_Inspect_pResultByColor-�ٽɵ���Ÿ ����Ʈ ���ν���
                        Set rs = oInspect.PrintResultByOrderExcel(sOrderID, nChkDate, sSDate, sEDate, _
                                            nChkLotNO, sLotNo, nChkExamNO, sExamNO)


                        Set oInspect = Nothing

                        nRow = 0
                        nCol = 0
                       ' iSQLFieldCnt = rs!FieldCount           'SQL���ν����� D,E,T,M �ʵ���� �տ� �������� �ʵ��� ��
                        
                        '****Excel PageFooter Start========================================================================================
                        Set oInspect = New PlusLib2.CInspect
                        oInspect.Connection = g_adoCon
                        ' �ҷ�����
                        'xp_Code_sDefectByLang->�� �ҷ����� ����Ʈ ��������(1:��,2:����)
                        Set rsDefect = oInspect.GetDefectByLang(IIf(optPrint(0), 1, 2))
                        Set oInspect = Nothing

                        '�ҷ����� �� ����Ʈ FOR��-�������� �ϴ� �߰����� �κ�
                        For k3 = 1 To rsDefect.RecordCount
                            If k3 > 49 Then Exit For            'S_201412_����_04 �� ���� ���� (OLD :i > 42)

                            ' S_201302_����_01 �� ���� ���� (OLD:58 + nRow, 3 + nCol)
                            'S_201412_����_04 �� ���� ���� (OLD : .Cells(58 + nRow, 2 + nCol)
                            .Cells(56 + nRow, 2 + nCol) = CStr(rsDefect!Tag) & "-" & CStr(CheckNull(rsDefect!Display))

                            ' S_201302_����_01 �� ���� ���� (OLD:k3 Mod 5)
                            If (k3 Mod 7) = 0 Then  '�ҷ������� 7�� ������ �� ��ȯ
                                nRow = nRow + 1
                                nCol = 0
                            Else
                                ' S_201302_����_01 �� ���� ����(OLD:1)
                                nCol = nCol + 11
                            End If

                            rsDefect.MoveNext
                        Next k3
                        Set rsDefect = Nothing
                        '****Excel PageFooter End========================================================================================

                        '����ġ ���ġ �˻����ǵ�..
                        '�˻�����
                        '�˻����� �ʱ�ȭ
                        sCondition = ""
                        If chkSearch(0).Value = vbChecked Then
                            If dtpDate(0) = dtpDate(1) Then
                                If optPrint(0).Value = True Then
                                    sDate = "�˻����� : " & MakeDate(DF_LONG, dtpDate(0))
                                  ElseIf optPrint(1).Value = True Then
                                    sDate = "INSPECT DATE : " & MakeDate(DF_LONG, dtpDate(0))
                               
                                End If
                            Else
                                If optPrint(0).Value = True Then
                                    sDate = "�˻����� : " & MakeDate(DF_LONG, dtpDate(0)) & "~" & MakeDate(DF_LONG, dtpDate(1))
                                ElseIf optPrint(1).Value = True Then
                                    sDate = "INSPECT DATE : " & MakeDate(DF_LONG, dtpDate(0)) & "~" & MakeDate(DF_LONG, dtpDate(1))
                               
                                End If
                            End If
                        Else
                            sDate = ""
                        End If
                        
                        sCondition = sCondition & sDate
                        
                        ' Lot ��ȣ
                        If chkSearch(2).Value = vbChecked Then
                            sLot = IIf(Len(sCondition) > 0, Space(5), "") & "LOT : " & txtSearch(2)
                        Else
                            sLot = ""
                        
                        End If
                        
                        sCondition = sCondition & sLot
                        
                            '��â���̿����� Roll ���þ����Ƿ� �ּ�
'''''                        ' ����ȣ
'''''                        If chkSearch(3).Value = vbChecked Then
'''''                            sRoll = IIf(Len(sCondition) > 0, Space(5), "")
'''''                                If optPrint(0).Value Then
'''''                                    sRoll = "����ȣ : "
'''''                                ElseIf optPrint(1).Value Then
'''''                                    sRoll = "RollNO : "
'''''
'''''                                End If
'''''                                sRoll = sRoll & txtSearch(4) & " ~ " & txtSearch(5)
'''''                        Else
'''''                            sRoll = ""
'''''                        End If
'''''
'''''                        sCondition = sCondition & sRoll

                        ' ȣ��
                        If chkSearch(3).Value = vbChecked Then
                            If optPrint(0).Value Then
                                sMachineNO = IIf(Len(sCondition) > 0, Space(5), "") & "ȣ�� : " & cboExamNo
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
                        
 
          
                        'S_201203_����_04 �� ���� ���� - NEW�ҽ�
                        .Cells(3, 73) = Format(Now, "YYYY/MM/DD")           '��������
                        .Cells(4, 9) = MakeOrderID(sOrderID, OM_EXPAND)          ' ������ȣ
                        .Cells(4, 30) = grdData.TextMatrix(grdData.Row, IIf(optPrint(0), grdData.ColIndex("Kcustom"), grdData.ColIndex("ECustom")))   '�ŷ�ó (4:�ѱ�,12:����)
                        .Cells(4, 53) = grdData.TextMatrix(grdData.Row, 7)                '���ַ�
                        .Cells(4, 71) = IIf(grdData.TextMatrix(grdData.Row, grdData.ColIndex("UnitClss")) = "0", "YARD", "METER")  '����
                       
                        .Cells(5, 9) = sOrderNO                                     'ORDER NO
                        .Cells(5, 30) = grdData.TextMatrix(grdData.Row, 4)          ' ǰ��
                        .Cells(5, 53) = grdData.TextMatrix(grdData.Row, 6)         ' ������
                        .Cells(5, 71) = grdData.TextMatrix(grdData.Row, 12)        ' ������
                        .Cells(6, 9) = Format(grdData.TextMatrix(grdData.Row, 9), "#,##0") & IIf(grdData.TextMatrix(grdData.Row, grdData.ColIndex("UnitClss")) = "0", "Y", "M")     ' �Ѽ���
                        .Cells(6, 27) = Format(grdData.TextMatrix(grdData.Row, 13), "#,##0.0")  ' �Ѻ���
                        .Cells(6, 44) = Format(grdData.TextMatrix(grdData.Row, 8), "#,##0")  '�� ����
                        .Cells(6, 62) = Format(grdData.TextMatrix(grdData.Row, 14), "#,##0")  '�� �ߺ�
                        .Cells(6, 77) = Format(grdData.TextMatrix(grdData.Row, 15), "#,##0")  '����
                        
                        
                        ' ��޺� ����
                        Set oInspect = New PlusLib2.CInspect
                        oInspect.Connection = g_adoCon
'                        Set rsGradeCnt = oInspect.GetGradeQtyByColor(sOrderID, sColorID, nChkDate, sSDate, sEDate, _
'                            nChkLotNO, sLotNo, nChkExamNO, sExamNO)
'
                        Set rsGradeCnt = oInspect.GetGradeQtyByOrder(sOrderID, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                            IIf(chkSearch(2), 1, 0), txtSearch(2), IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))

                        '�հ�/���հ� ����-�� ��޺� ���� A-D�ۿ� ����
                        For iLoop = 0 To rsGradeCnt.RecordCount - 1

                            If rsGradeCnt!GradeID = "1" Then               '�հ�-A�
                                .Cells(7, 11) = IIf(IsNull(rsGradeCnt!GradeCount), "0", Format(rsGradeCnt!GradeCount, "#,##0"))
                            ElseIf rsGradeCnt!GradeID = "2" Then           '�ҷ� B���
                                .Cells(7, 24) = IIf(IsNull(rsGradeCnt!GradeCount), "0", Format(rsGradeCnt!GradeCount, "#,##0"))
                            ElseIf rsGradeCnt!GradeID = "3" Then           '�ҷ� C���
                                .Cells(7, 36) = IIf(IsNull(rsGradeCnt!GradeCount), "0", Format(rsGradeCnt!GradeCount, "#,##0"))
                            ElseIf rsGradeCnt!GradeID = "4" Then           '�ҷ� F���
''                                .Cells(7, 49) = IIf(IsNull(rsGradeCnt!GradeCount), "0", Format(rsGradeCnt!GradeCount, "#,##0"))         'D�����ġ
                                .Cells(7, 73) = IIf(IsNull(rsGradeCnt!GradeCount), "0", Format(rsGradeCnt!GradeCount, "#,##0"))           'F�����ġ
                            End If

                            rsGradeCnt.MoveNext
                        Next iLoop
                        rsGradeCnt.Close
                        Set rsGradeCnt = Nothing
                        
                        '****Excel PageHeader End========================================================================================
                        
                        nBaseRow = GetExcelBaseRow(nPage)
                        nRow = 0
                        nRecordCount = rs.RecordCount
                        
                        '������ �߰�
                        Call InsertExcelForm(xlApp, nPage)
                        .Worksheets("Report").Activate
                        For j3 = 0 To rs.RecordCount - 1
                        
                             If nRow > iExcelByPage - 2 Then            '46�� �̻��� ��� ������ ����(OLD : nRow > 46)
                                nPage = nPage + 1
                                Call InsertExcelForm(xlApp, nPage)
                                nBaseRow = GetExcelBaseRow(nPage)
                                nRow = 0
                             End If


 
   
                            lnCurRow = iDataStartRow + nBaseRow + nRow
                            '������ ���
                            If CheckNull(rs!Cls) = 1 Then
                                                             
                                .Range("A" & lnCurRow & ":BG" & lnCurRow + 1).Select
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
                                .Cells(lnCurRow, 1) = CheckNull(rs!Color)                'Color
                                
                                 
                            Else
                                 .Cells(lnCurRow, 1) = CheckNull(rs!LotNo)                'LOT No
                                'S_201412_����_04 �� ���� ����( 9 ��� iDataStartRow, 10��� iDataStartRow+1 ���)
                                .Cells(lnCurRow, 4) = CheckNull(rs!RollNo)             '����ȣ
                                '.Cells(iDataStartRow + nBaseRow + nRow, 7) = rs!ExamNO          '�˻�ȣ��
                                .Cells(lnCurRow, 7) = CheckNull(rs!Person)              '�˻��� ��� ExamNo
                                
                                If Trim(CheckNull(rs!GradeID)) = "" Then        '���
                                    .Cells(lnCurRow, 73) = ""
                                Else
                                    If rs!GradeID <> 4 Then
                                        .Cells(lnCurRow, 73) = Chr(Asc("@") + CheckNull(CInt(rs!GradeID)))          'ABC(1~3) ���
                                    Else
                                        .Cells(lnCurRow, 73) = "F"                                       'F���(4)
                                    End If
                                End If
    
                                '��ǥ�ҷ�
                                .Cells(lnCurRow, 76) = IIf(optPrint(0), CheckNull(rs!DefectKor), CheckNull(rs!DefectEng))
                                
                                
    
                                iExcelCol = 0
                                iInsRow = 0
    
    
                                nDefectQty = CheckNum(rs!DefectQty)         '�� �ҷ� ����
                                iExcelCol = 0
    
    
    
                                For iDefCnt = 1 To nDefectQty        'D,E,T,M (4�� ������ ������)
    
                                    If nRow > iExcelByPage - 2 Then            '44(iExcelByPage - 2) �� ������ ����(OLD : nRow > 46)
                                       nPage = nPage + 1
                                       Call InsertExcelForm(xlApp, nPage)
                                       nBaseRow = GetExcelBaseRow(nPage)
                                       nRow = 0
                                    End If
    
    ''                                '//�ҷ� ����
                                    'S_201412_����_04 �� ���� ����( 9 ��� iDataStartRow, 10��� iDataStartRow+1 ���)
                                    '�׸� ���ۿ��� iExcelDefStartCol(14)�� - 4���� �⺻��Ʈ �̹Ƿ� X 4(iDefectColRep)
                            
                                    .Cells(lnCurRow, iExcelDefStartCol + iExcelCol) = Format(CheckNull(rs(iSQLFieldCnt + (iDefCnt))), "#,##0")    '�ҷ���ġ
    
    ''                              '�ҷ��Ʒ���
                                    'S_201901_��������_02 �� ���� ���� : 30 -> 60
                                    .Cells(lnCurRow + 1, iExcelDefStartCol + iExcelCol) = CheckNull(rs(iSQLFieldCnt + 60 + (iDefCnt)))      '�ҷ��±�
 
    
                                    If (iExcelDefStartCol + iExcelCol) >= iExcelDefStartCol + (iDefectCntByRow - 1) * iDefectColSpan Then         '1 �� �μ� �Ϸ�Ǹ� �����ٷ� ��� �μ�(�⺻ �ҷ� ���ٿ� 15��) (14������ 56��)
                                        iExcelCol = 0
    
                                        If nDefectQty Mod 15 <> 0 Or iDefCnt <> nDefectQty Then                      '15�� ����� �������� ������ ���� ��Ű�� ����
                                            nRow = nRow + 2     '���� ������ �̷��Ը� ����
                                        End If
    ''                                    iTimes = iTimes - 1       '//�ҷ��� 15���� ���ΰ��(15,30,45,60 ���� ��� ������ �� ����� ���� üũ����)
                                    Else
    '                                    iExcelCol = iExcelCol + 3
                                        iExcelCol = iExcelCol + iDefectColSpan
                                    End If
    
                                Next iDefCnt
                            End If
                               
                               
                               
                            .Cells(lnCurRow, 10) = CheckNull(rs!StuffQty)                                '���Է�

                            .Cells(lnCurRow, 59) = CheckNull(Format(rs!SampleQty, "0"))                 '�ߺ�
                            .Cells(lnCurRow, 62) = CheckNull(Format(rs!CutQty, "0"))                     '����
                            .Cells(lnCurRow, 64) = CheckNull(Format(rs!CtrlQty, "#,##0"))                 '�Ǽ���

 
                            .Cells(lnCurRow, 68) = CheckNull(Format(rs!LossQty, "0"))                    '����
                            .Cells(lnCurRow, 70) = CheckNull(rs!DefectQty)                               '�ҷ�����
 

'
                            nStuffQty = nStuffQty + rs!StuffQty                 '�� ���Է�
                            nWeight = nWeight + rs!CtrlQty                      '�� �߷�



                            nRow = nRow + 2

                            rs.MoveNext
                        Next j3
                        
                        rs.Close

                        'nRow 44 �ʰ��ϰ�� ������ �߰�
                        If nRow > iExcelByPage - 2 Then            '44(iExcelByPage - 2) �� ������ ����(OLD : nRow > 46)
                            nPage = nPage + 1
                            Call InsertExcelForm(xlApp, nPage)
                            nBaseRow = GetExcelBaseRow(nPage)
                            nRow = 0
                        End If
                        Call ExcelTotal(xlApp, nPage, nBaseRow, nRow)
                                
                        If optPrint(0).Value Then
                            .Cells(iDataStartRow + nBaseRow + nRow, 1) = "�հ� : "
                        ElseIf optPrint(1).Value Then
                            .Cells(iDataStartRow + nBaseRow + nRow, 1) = "TOTAL : "

                        End If
                        .Cells(iDataStartRow + nBaseRow + nRow, 7) = Format(nStuffQty, "#,###")         '�� ���Է�
                        
                        '//�Ʒ� �� �ߺ�, �� ����,�� �Ǽ���, �� ���� ������ ����κп��� ��µ�
''
                    End With            '// With xlApp �� ��
                
                nPage = nPage + 1   '���� �ҽ��� ����-Į�� �ٲ� ������ �߰�
                    
         '   End If
  '      Next i
    End With        '//With grdData�� ��
    
    
    Set oFs = Nothing

    Call xlBook.SaveAs(sReport)

    '//�̸����� ���üũ**********
    If Not bDirect Then         '�̸����� ����
        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1

        Call ProcessClose("XLMAIN")
    Else                        '�̸����� ����
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
''    Resume Next

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
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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

    Set rs = oInspect.GetResultByOrder(IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(1), IIf(optOrder(0), 2, 1), 0), IIf(optOrder(0), txtSearch(1), MakeOrderID(txtSearch(1), OM_REDUCE)), _
        IIf(chkSearch(2), 1, 0), txtSearch(2), IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    Set oInspect = Nothing

    With grdData
        .Redraw = False

        iNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
            .AddItem CStr(i) & vbTab & rs!OrderNo & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                rs!kCustom & vbTab & CStr(rs!Article) & vbTab & MakeDate(DF_LONG, rs!DvlyDate) & vbTab & _
                rs!WorkName & vbTab & CheckNum(rs!OrderQty) & vbTab & CheckNum(rs!PassRoll) & vbTab & _
                CheckNum(rs!PassQty) & vbTab & CheckNull(rs!ECustom) & vbTab & rs!UnitClss & vbTab & _
                rs!WorkWidth & vbTab & CheckNum(rs!LossQty) & vbTab & CheckNum(rs!SampleQty) & vbTab & _
                CheckNum(rs!CutQty)

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

        .Redraw = True
        .SetFocus
    End With

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        chkSearch(1).Caption = "Order No"
    Else
        chkSearch(1).Caption = "������ȣ"
    End If
    cmdSearch.SetFocus
End Sub

Private Sub optPrint_Click(Index As Integer)
    cmdPrint.SetFocus
End Sub

Private Sub cmdExcel_Click()
    Call MakeExcelGrid(grdData)
End Sub

Private Sub cmdPrint_Click()
    Dim oInspect As PlusLib2.CInspect
    Dim rs       As ADODB.Recordset
    Dim sParam() As String
    Dim iPoint%, iLoop%, i%
    Dim sOrderID$

    On Error GoTo ErrHandler

    Me.PopupMenu PlusMDI.mnuPopup

    Screen.MousePointer = vbHourglass

    With grdData
        sOrderID = MakeOrderID(.TextMatrix(.Row, 2), OM_REDUCE)
    End With

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon

    Set rs = oInspect.GetDefectByLang(IIf(optPrint(0), 1, 2))
    Set oInspect = Nothing

    ' footer
    ReDim sParam(39)
    For i = 0 To rs.RecordCount - 1
        sParam(i) = CStr(rs!Tag) & "-" & CStr(CheckNull(rs!Display))

        rs.MoveNext
    Next i
    rs.Close
    Set rs = Nothing

    Do While i <= 39
        sParam(i) = " "
        i = i + 1
    Loop

    ReDim Preserve sParam(53)
    ' MarkClss
    If optPrint(0) Then
        If dtpDate(0) = dtpDate(1) Then
            sParam(i) = "�˻����� : " & Format(dtpDate(0), "YYYY�� MM�� DD��") & Space(5) & _
                IIf(chkSearch(3), "ȣ�� : " & cboExamNo, "��ü ȣ��")
        Else
            sParam(i) = "�˻����� : " & Format(dtpDate(0), "YYYY�� MM�� DD��") & " ~ " & Format(dtpDate(1), "YYYY�� MM�� DD��") & Space(5) & _
                IIf(chkSearch(3), "ȣ�� : " & cboExamNo, "��ü ȣ��")
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
    sParam(i + 1) = grdData.TextArray(iPoint + IIf(optPrint(0), 3, 10))
    sParam(i + 2) = grdData.TextArray(iPoint + 7)
    sParam(i + 3) = IIf(grdData.TextArray(iPoint + 11) = "0", "YARD", "METER")
    sParam(i + 4) = grdData.TextArray(iPoint + 1)
    sParam(i + 5) = grdData.TextArray(iPoint + 4)
    sParam(i + 6) = IIf(grdData.TextArray(iPoint + 6) = "", " ", grdData.TextArray(iPoint + 6))
    sParam(i + 7) = grdData.TextArray(iPoint + 12)
    sParam(i + 8) = Format(grdData.TextArray(iPoint + 9), "#,##0") & IIf(grdData.TextArray(iPoint + 11) = "0", "Y", "M")
    sParam(i + 9) = Format(grdData.TextArray(iPoint + 13), "#,##0")
    sParam(i + 10) = grdData.TextArray(iPoint + 8)
    sParam(i + 11) = Format(grdData.TextArray(iPoint + 14), "#,##0")
    sParam(i + 12) = Format(grdData.TextArray(iPoint + 15), "#,##0")
    i = i + 13

    ReDim Preserve sParam(61)
    ' title
    If optPrint(0) Then
        sParam(i) = "�˻� ���ǥ"
    Else
        sParam(i) = "INSPECTION REPORT"
    End If
    ' companyname
    sParam(i + 1) = CompanyName
    i = i + 2

    Set oInspect = New PlusLib2.CInspect
    oInspect.Connection = g_adoCon
   
    Set rs = oInspect.GetGradeQtyByOrder(sOrderID, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(2), 1, 0), txtSearch(2), IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    
    For iLoop = 0 To rs.RecordCount - 1
        sParam(i + iLoop) = IIf(IsNull(rs!GradeCount), "0", rs!GradeCount)

        rs.MoveNext
    Next iLoop
    rs.Close
    Set rs = Nothing
    If optPrint(0) Then
    Set rs = oInspect.PrintResultByOrder(sOrderID, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
        IIf(chkSearch(2), 1, 0), txtSearch(2), IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    Else: Set rs = oInspect.PrintResultByOrder_eng(sOrderID, IIf(chkSearch(0), 1, 0), MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)) _
        , IIf(chkSearch(2), 1, 0), txtSearch(2), IIf(chkSearch(3), 1, 0), Format(Left(cboExamNo, 2), "00"))
    End If
    Set oInspect = Nothing

    Call PrintReport(IIf(optPrint(0), REPORTFILE_1, REPORTFILE_2), rs, sParam, PlusMDI.PrintPreview)

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oInspect = Nothing

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    With grdData
        .Cols = 16
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .TextArray(1) = "Order No.":    .ColWidth(1) = 1540:    .ColAlignment(1) = flexAlignLeftCenter:     .ColKey(1) = "OrderNo"
        .TextArray(2) = "������ȣ":     .ColWidth(2) = 1225:    .ColAlignment(2) = flexAlignCenterCenter:   .ColKey(2) = "OrderID"
        .TextArray(3) = "�ŷ�ó":       .ColWidth(3) = 1300:    .ColAlignment(3) = flexAlignLeftCenter:     .ColKey(3) = "Kcustom"
        .TextArray(4) = "ǰ��":         .ColWidth(4) = 2000:    .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "��������":     .ColWidth(5) = 990:     .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "��������":     .ColWidth(6) = 900:     .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "���ּ���":     .ColWidth(7) = 1010:    .ColAlignment(7) = flexAlignRightCenter:    .ColFormat(7) = GetFormat()
        .TextArray(8) = "�˻�����":     .ColWidth(8) = 900:     .ColAlignment(8) = flexAlignRightCenter:    .ColFormat(8) = GetFormat()
        .TextArray(9) = "�˻����":     .ColWidth(9) = 1110:    .ColAlignment(9) = flexAlignRightCenter:    .ColFormat(9) = GetFormat(g_nPointPos)
'        .TextArray(10) = "�ŷ�ó(��)":  .ColWidth(10) = 0:      .ColKey(10) = "UnitClss"
'        .TextArray(11) = "��������":    .ColWidth(11) = 0:      .ColKey(11) = "ECustom" 'S_201901_��������_04 �� ���� ���� : ColKey �߰���
        
        .TextArray(10) = "�ŷ�ó(��)":  .ColWidth(10) = 0:      .ColKey(10) = "ECustom"
        .TextArray(11) = "��������":    .ColWidth(11) = 0:      .ColKey(11) = "UnitClss" 'S_201901_��������_04 �� ���� ���� : ColKey �߰���
        .TextArray(12) = "������":      .ColWidth(12) = 0
        .TextArray(13) = "�������":    .ColWidth(13) = 0:      .ColFormat(13) = GetFormat(1)
        .TextArray(14) = "�ߺ�����":    .ColWidth(14) = 0:      .ColFormat(14) = GetFormat(g_nPointPos)
        .TextArray(15) = "���ܼ���":    .ColWidth(15) = 0:      .ColFormat(15) = GetFormat(g_nPointPos)

        .Redraw = flexRDDirect
    End With
End Sub

'S_201211_��������_03 �� ���� �߰�
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
'2011.03.13 �߰�
Private Function ExcelTotal(xlApp As Excel.Application, nPage As Integer, nBaseRow As Integer, nRow As Integer)

    On Error GoTo Err_Rtn
    
    With xlApp
        'S_201302_����_01 �� ���� ����-OLD�ҽ�
        '.Range("A" & 9 + nBaseRow + nRow & ":F" & 9 + nBaseRow + 47).Select
        
        'S_201302_����_01 �� ���� ����-NEW�ҽ�
        'iDataStartRow + nBaseRow + nRow => Merge �� From ��
        'iDataStartRow + nBaseRow + iExcelByPage - 1   => Merge �� To ��
''        .Range(GF_Excel_CA(1) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(6) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '���ڸ� Excel Column �����ڷ� ���� Range ����
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 1), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 6)).Select         'R1C1 ��Ÿ�� ���� �ּ�
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

        'S_201302_����_01 �� ���� ����-OLD�ҽ�
        '.Range("G" & 9 + nBaseRow + nRow & ":M" & 9 + nBaseRow + 47).Select
        'S_201302_����_01 �� ���� ����-NEW�ҽ�
''        .Range(GF_Excel_CA(7) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(13) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '���ڸ� Excel Column �����ڷ� ���� Range ����
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 7), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 13)).Select         'R1C1 ��Ÿ�� ���� �ּ�
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

        'S_201302_����_01 �� ���� ����-OLD�ҽ�
        '.Range("N" & 9 + nBaseRow + nRow & ":BF" & 9 + nBaseRow + 47).Select
        'S_201302_����_01 �� ���� ����-NEW�ҽ�
''        .Range(GF_Excel_CA(14) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(58) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select '���ڸ� Excel Column �����ڷ� ���� Range ����
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 14), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 58)).Select         'R1C1 ��Ÿ�� ���� �ּ�
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
        
         
        'S_201302_����_01 �� ���� ����-OLD�ҽ�
        '.Range("BG" & 9 + nBaseRow + nRow & ":BK" & 9 + nBaseRow + 47).Select
        'S_201302_����_01 �� ���� ����-NEW�ҽ�
''        .Range(GF_Excel_CA(59) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(63) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '���ڸ� Excel Column �����ڷ� ���� Range ����
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 59), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 63)).Select         'R1C1 ��Ÿ�� ���� �ּ�
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

        
        'S_201302_����_01 �� ���� ����-OLD�ҽ�
        '.Range("BL" & 9 + nBaseRow + nRow & ":BO" & 9 + nBaseRow + 47).Select
        'S_201302_����_01 �� ���� ����-NEW�ҽ�
''        .Range(GF_Excel_CA(64) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(67) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '���ڸ� Excel Column �����ڷ� ���� Range ����
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 64), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 67)).Select         'R1C1 ��Ÿ�� ���� �ּ�
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
               
        'S_201302_����_01 �� ���� ����-OLD�ҽ�
        '.Range("BP" & 9 + nBaseRow + nRow & ":BT" & 9 + nBaseRow + 47).Select
        'S_201302_����_01 �� ���� ����-NEW�ҽ�
''        .Range(GF_Excel_CA(68) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(72) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '���ڸ� Excel Column �����ڷ� ���� Range ����
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 68), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 72)).Select         'R1C1 ��Ÿ�� ���� �ּ�
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
        
        
        'S_201302_����_01 �� ���� ����-OLD�ҽ�
        '.Range("BU" & 9 + nBaseRow + nRow & ":CF" & 9 + nBaseRow + 47).Select
        'S_201302_����_01 �� ���� ����-NEW�ҽ�
''        .Range(GF_Excel_CA(73) & (iDataStartRow + nBaseRow + nRow), GF_Excel_CA(84) & (iDataStartRow + nBaseRow + iExcelByPage - 1)).Select  '���ڸ� Excel Column �����ڷ� ���� Range ����
        .Range(.Cells(iDataStartRow + nBaseRow + nRow, 73), .Cells(iDataStartRow + nBaseRow + iExcelByPage - 1, 84)).Select         'R1C1 ��Ÿ�� ���� �ּ�
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
