VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProcCostReport 
   Caption         =   "Order��û����"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   13485
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   6990
      TabIndex        =   30
      Top             =   8490
      Width           =   1425
      Begin VB.OptionButton optGub 
         Caption         =   "�������"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   32
         Top             =   120
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optGub 
         Caption         =   "��ü����"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   31
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.ComboBox cboTaxClss 
      Height          =   300
      Left            =   1350
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   24
      Top             =   360
      Width           =   1395
   End
   Begin VB.ComboBox CboOrderFlag 
      Height          =   300
      Left            =   10695
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   22
      Top             =   150
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Threed.SSPanel pnlPrint 
      Height          =   3195
      Left            =   3720
      TabIndex        =   11
      Top             =   2010
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   5636
      _Version        =   196610
      BevelWidth      =   3
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel SSPanel6 
         Height          =   405
         Left            =   90
         TabIndex        =   18
         Top             =   90
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   714
         _Version        =   196610
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "Order�� û���� �μ�"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   2070
         TabIndex        =   14
         Top             =   630
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1296
         _Version        =   196610
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optPrn 
            Caption         =   "��ü�� û����"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   16
            Top             =   420
            Width           =   1605
         End
         Begin VB.OptionButton optPrn 
            Caption         =   "��ü��Ȳ"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   15
            Top             =   120
            Value           =   -1  'True
            Width           =   1605
         End
      End
      Begin VB.ComboBox cboCustom 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2070
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   13
         Top             =   1410
         Width           =   2040
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   315
         Left            =   780
         TabIndex        =   12
         Top             =   1410
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "�μ����"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Index           =   0
         Left            =   780
         TabIndex        =   17
         Top             =   630
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "�μⱸ��"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdPrnCancel 
         Height          =   495
         Left            =   3540
         TabIndex        =   19
         Top             =   2430
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196610
         Caption         =   "���"
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   2070
         TabIndex        =   20
         Top             =   1770
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyy�� MM�� dd��"
         Format          =   115802115
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Index           =   1
         Left            =   780
         TabIndex        =   21
         Top             =   1770
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "�μ�����"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdOrderPrint_OLD 
         Height          =   495
         Left            =   1980
         TabIndex        =   26
         Top             =   2430
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196610
         Caption         =   "�������μ�"
      End
      Begin Threed.SSCommand cmdPrnOK 
         Height          =   495
         Left            =   420
         TabIndex        =   29
         Top             =   2430
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196610
         Caption         =   "�μ�"
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   3075
         Left            =   60
         Top             =   60
         Width           =   5265
      End
   End
   Begin VB.TextBox txtCustom 
      Height          =   300
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   30
      Width           =   2025
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "�˻�(&F)"
      Height          =   630
      Left            =   7110
      MousePointer    =   99  '����� ����
      Style           =   1  '�׷���
      TabIndex        =   3
      ToolTipText     =   "�ڷ� ����"
      Top             =   30
      Width           =   780
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Index           =   0
      Left            =   1350
      TabIndex        =   0
      Top             =   30
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyy�� MM�� dd��"
      Format          =   115802115
      CurrentDate     =   36871
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   6
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196610
      Caption         =   "û�����"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   11820
      TabIndex        =   5
      Top             =   8520
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      �ݱ�(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7800
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   660
      Width           =   13440
      _cx             =   23707
      _cy             =   13758
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
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
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmProcCostReport.frx":0000
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
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8460
      TabIndex        =   7
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "û���� �μ�(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   285
      Index           =   0
      Left            =   6720
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      _Version        =   196610
      ButtonStyle     =   3
      Outline         =   0   'False
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   9
      Left            =   3390
      TabIndex        =   9
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   196610
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkSearch 
         Caption         =   "�� �� ó"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   975
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   510
      Index           =   1
      Left            =   30
      TabIndex        =   10
      Top             =   8520
      Visible         =   0   'False
      Width           =   2100
      _cx             =   3704
      _cy             =   900
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   16777215
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
      Index           =   0
      Left            =   9360
      TabIndex        =   23
      Top             =   150
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196610
      Caption         =   "��±���"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlCaption 
      Height          =   300
      Index           =   1
      Left            =   30
      TabIndex        =   25
      Top             =   360
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   196610
      Caption         =   "��뱸��"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VSFlex7LCtl.VSFlexGrid grdDataOrder 
      Height          =   600
      Left            =   930
      TabIndex        =   27
      Top             =   8580
      Visible         =   0   'False
      Width           =   4110
      _cx             =   7250
      _cy             =   1058
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
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
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmProcCostReport.frx":0131
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
      Left            =   5310
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "      ����(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdOrderPrint 
      Height          =   690
      Left            =   10140
      TabIndex        =   33
      Top             =   8520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196610
      Caption         =   "������û���� �μ�"
   End
End
Attribute VB_Name = "frmProcCostReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' �����̷�
'------------------------------------------------------------------------------
'��û���� : 2012.02.13
'��û���� : û���� ����� ��Ϲ�ȣ �߸���(503-41-73761)
'S_201202_��������_01 �� ���� ����
'���泻�� : ����� ��Ϲ�ȣ ����
'
'��ûID : S_201203_��������_02
'��û���� : 2012.03.05
'��û���� : ������ �� ��µǰ�
'���泻�� : cmdOrderPrint ��ư �߰�
'
'��ûID : S_201211_��������_02
'��û���� : 2012.11.22
'��û���� : ǰ�� û������ ������ ��µǰ�
'���泻�� : ���� ������� ����(���� �׸��� ���)
'
'��ûID : S_201212_��������_06
'��û���� : 2012.12.20
'��û���� : û����-������ڻ���, �ϴܿ� ������ ǥ��
'           ������û����-������ڻ���, �ϴܿ� ������ ǥ��, ������ ����ǥ��
'���泻�� :
'
'2013.12.12   ��ü    ���¿�   S_201312_��������_99   �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�,�ŷ�ó �ּ� ���θ� �ּ� Select
'******************************************************************************
Option Explicit

Dim sPrinter As String

'S_201203_��������_02 �� ���� �߰�
Private Const EXCEL_ROW As Integer = 42             '���� �� ������ �� ���(����Ʈ ���� ��)

'û����-������
Private Const Reportfile_Excel_Order = "\Report\ProcCostReportByOrder.xls"    'S_201211_��������_02 �� ���� ����(OLD:ProcCostReport.xls)

'S_201211_��������_02 �� ���� �߰�
'û����-ǰ��
Private Const Reportfile_Excel = "\Report\ProcCostReport.xls"

Private Sub chkSearch_Click(Index As Integer)
    Select Case Index

        Case 1    '�ŷ�ó
            If chkSearch(Index) = vbChecked Then
                txtCustom(1).Enabled = True
                txtCustom(1).SetFocus
                cmdFind(0).Enabled = True
            Else
                txtCustom(1).Enabled = False
                cmdFind(0).Enabled = False
                txtCustom(1).Tag = ""
            End If
            
    End Select
End Sub

Private Sub chkSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub cmdExcel_Click()
    If grdDataOrder.Rows = grdDataOrder.FixedRows Then Exit Sub

    Call MakeExcelGrid(grdDataOrder)

End Sub

'Private Sub chkSearch_Click()
'    If chkSearch.Value = vbChecked Then
'        dtpDate(0).Enabled = True
'        dtpDate(1).Enabled = True
'    Else
'        dtpDate(0).Enabled = False
'        dtpDate(1).Enabled = False
'    End If
'End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0                '[1] �ŷ�ó �ڵ�
            Call ReturnCode(LG_CUSTOM, , False, txtCustom(1))
    End Select
End Sub

'S_201203_��������_02 �� ���� �߰�
Private Sub cmdOrderPrint_Click()

    'S_201211_��������_02 �� ���� OLD�ҽ�
''    If optPrn(0).Value = True Or cboCustom.ListIndex <= 0 Then
''        MsgBox "������ ��ü ��Ȳ�� ����Ҽ� �����ϴ�.", vbOKOnly, "��� �Ұ�"
''        Exit Sub
''    Else
''
''        Call ExcelPrintByOneCustByOrder(PlusMDI.PrintPreview)         '1���� ���þ�ü ���
''''        '������ ���� ���
''''        Call SetDataToPrn(cboCustom.Text)
''
''''        Call ReturnPrinter(sPrinter)
''        pnlPrint.Visible = False
''    End If

    'S_201211_��������_02 �� ���� NEW �ҽ�
    cmdExcel.Enabled = False
    Frame1.Enabled = False
    cmdPrint.Enabled = False
    cmdOrderPrint.Enabled = False
    cmdExit.Enabled = False
    
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass
    
    If optGub(0).Value = True Then
        If Trim(txtCustom(1).Tag) = "" Then
            If grdData(0).FixedRows >= grdData(0).Row Or grdData(0).TextMatrix(grdData(0).Row, 9) = "" Then
                Screen.MousePointer = vbDefault
                cmdExcel.Enabled = True
                Frame1.Enabled = True
                cmdPrint.Enabled = True
                cmdOrderPrint.Enabled = True
                cmdExit.Enabled = True
    
                MsgBox "�ŷ�ó�� ������ �� �μ��Ͻʽÿ�.", vbOKOnly, "��� �Ұ�"
                Exit Sub
            End If

        End If
        Call ExcelPrintByOneCustByOrder(PlusMDI.PrintPreview)         '1���� ���þ�ü ���
    Else

    
        MsgBox "������ ��ü ��Ȳ�� ����Ҽ� �����ϴ�.", vbOKOnly, "��� �Ұ�"

    End If
    
    Screen.MousePointer = vbDefault
    cmdExcel.Enabled = True
    Frame1.Enabled = True
    cmdPrint.Enabled = True
    cmdOrderPrint.Enabled = True
    cmdExit.Enabled = True
    
    
End Sub

'S_201211_��������_02 �� ���� ����-NEW �ҽ�
Private Sub cmdPrint_Click()
    
    cmdExcel.Enabled = False
    Frame1.Enabled = False
    cmdPrint.Enabled = False
    cmdOrderPrint.Enabled = False
    cmdExit.Enabled = False
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Screen.MousePointer = vbHourglass
    
    'S_201211_��������_02 �� ���� �߰�--------------------------------------------------------------------------
    If optGub(0).Value = True Then
        If Trim(txtCustom(1).Tag) = "" Then

    
            If grdData(0).FixedRows >= grdData(0).Row Or grdData(0).TextMatrix(grdData(0).Row, 9) = "" Then
                Screen.MousePointer = vbDefault
                cmdExcel.Enabled = True
                Frame1.Enabled = True
                cmdPrint.Enabled = True
                cmdOrderPrint.Enabled = True
                cmdExit.Enabled = True
                
                MsgBox "�ŷ�ó�� ������ �� �μ��Ͻʽÿ�.", vbOKOnly, "��ºҰ�"
                Exit Sub
            End If

        End If
        Call ExcelPrintByOneCust(PlusMDI.PrintPreview)         '1���� ���þ�ü ���
    Else
        Call ExcelPrintByAllCust(PlusMDI.PrintPreview)        '��ü��ü ���
    End If
    '------------------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    cmdExcel.Enabled = True
    Frame1.Enabled = True
    cmdPrint.Enabled = True
    cmdOrderPrint.Enabled = True
    cmdExit.Enabled = True
End Sub



'S_201211_��������_02 �� ���� ����-�ּ�ó��
''Private Sub cmdPrint_Click()
''    pnlPrint.Visible = True
''End Sub

Private Sub cmdPrnCancel_Click()
    pnlPrint.Visible = False
End Sub


'S_201211_��������_02 �� ���� ����-�ּ�ó��
''Private Sub cmdPrnOK_Click()
''    Dim II%
''
''    If optPrn(0).Value = True Then
''        Call FillGrdList
''    Else
''        'S_201203_��������_02 �� ���� ����
'''        If cboCustom.Text = AllStr Then         '��ü�����̸�
''        If cboCustom.ListIndex = 0 Then
''            For II = 1 To cboCustom.ListCount - 1
''                'S_201203_��������_02 �� ���� ����
''''                Call SetDataToPrn(cboCustom.Text)
''                Call SetDataToPrn(RTrim(Left(cboCustom.List(cboCustom.ListIndex), 50)))
''            Next II
''        Else
''            'S_201203_��������_02 �� ���� ����
''''            Call SetDataToPrn(cboCustom.Text)
''            Call SetDataToPrn(RTrim(Left(cboCustom.List(cboCustom.ListIndex), 50)))
''        End If
''    End If
''
''    Call ReturnPrinter(sPrinter)
''    pnlPrint.Visible = False
''
''End Sub

Private Sub cmdSearch_Click()
    Call FillgrdData
End Sub

Sub FillGrdList()
    Dim i%, II%, JJ%
    Dim sDate As String, eDate As String, nRows As Integer
    
    sDate = MakeDate(DF_MD, dtpDate(0))
       
    With grdData(0)
        
        Call SetPrintMode(grdData(0), 1, True)
        
        nRows = 1
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "ORDER�� û���� ��Ȳ"
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = True
        .Cell(flexcpAlignment, nRows, 0, nRows, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(nRows) = 800
        
        nRows = 2
        .RowHeight(nRows) = 500
        .Cell(flexcpText, nRows, 2, nRows, .Cols - 1) = "�� ������ : " & sDate
        .Cell(flexcpAlignment, nRows, 2, nRows, .Cols - 1) = flexAlignCenterCenter
        
        
        .RowHidden(3) = True
        .RowHidden(4) = True
        .RowHidden(5) = True
        
        .MergeCells = flexMergeFree
        For i = 0 To .FixedRows - 1
           .MergeRow(i) = True
        Next i

        .ExtendLastCol = False
        .PrintGrid "��������", True, 2, 700, 500
        
        Call SetPrintMode(grdData(0), 1, False)
        .ExtendLastCol = True
    End With

End Sub


Sub FillGrdPrintHeader(ByVal kCustom As String)
    Dim i%
    Dim sDate As String
    Dim nRows As Integer
    Dim eDate As String

    With grdData(1)
        .Rows = 7
        .Cols = 11
        .FixedRows = 7
        .RowHeightMin = 400
        
        nRows = 0
        .Cell(flexcpText, nRows, 2, nRows, .Cols - 1) = "ORDER�� û����"
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = True
        .Cell(flexcpFontUnderline, nRows, 0, nRows, .Cols - 1) = True
        .Cell(flexcpAlignment, nRows, 0, nRows, .Cols - 1) = flexAlignCenterCenter
        
        .RowHeight(nRows) = 800
        .RowHeight(3) = 400
        .RowHeight(4) = 400

        nRows = 1
        .RowHeight(nRows) = 500
        .Cell(flexcpText, nRows, 2, nRows, .Cols - 1) = "�� �� �� ó : " & kCustom
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 11
        .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = False
        .Cell(flexcpAlignment, nRows, 2, nRows, .Cols - 1) = flexAlignCenterCenter


        nRows = 2
        .RowHeight(2) = 500
        .Cell(flexcpText, nRows, 2, nRows, .Cols - 1) = "�� �������� : " & MakeDate(DF_FULL, dtpDate(0))
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 11
        
        
        'S_201202_��������_01 �� ���� ����( OLD: 504-41-73761)
        .Cell(flexcpText, 4, 2, 4, 3) = "��Ϲ�ȣ : 503-41-73761"
        .Cell(flexcpText, 5, 2, 5, 3) = "��    �� : �뱸 ���� ���7�� 2037-28"
        
        .Cell(flexcpText, 4, 4, 4, 6) = "�� ȣ : ��������"
        .Cell(flexcpText, 5, 4, 5, 6) = "�� �� : ������"
        
        .Cell(flexcpText, 4, 7, 4, 8) = "��  ǥ : �� �� ��"
        .Cell(flexcpText, 5, 7, 5, 8) = "��  �� : ��������"
        
        .Cell(flexcpAlignment, 4, 0, 4, .Cols - 1) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 5, 0, 5, .Cols - 1) = flexAlignLeftCenter

        .ColHidden(1) = True

        For i = 0 To 6
           .MergeRow(i) = True
        Next i

        nRows = 6
        .RowHeight(nRows) = 400
        
        .TextMatrix(nRows, 0) = "":                 .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "�ŷ�ó":           .ColWidth(1) = 0:          .ColAlignment(1) = flexAlignLeftCenter:       .FixedAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "ǰ��":             .ColWidth(2) = 3000:       .ColAlignment(2) = flexAlignLeftCenter:       .FixedAlignment(2) = flexAlignCenterCenter
        .TextMatrix(nRows, 3) = "��������":         .ColWidth(3) = 1800:       .ColAlignment(3) = flexAlignCenterCenter:     .FixedAlignment(3) = flexAlignCenterCenter
        .TextMatrix(nRows, 4) = "��    ��":         .ColWidth(4) = 1600:       .ColAlignment(4) = flexAlignRightCenter:      .FixedAlignment(4) = flexAlignCenterCenter
        
        .TextMatrix(nRows, 5) = "�ܰ�":             .ColWidth(5) = 1200:       .ColAlignment(5) = flexAlignRightCenter:      .FixedAlignment(5) = flexAlignCenterCenter
        .TextMatrix(nRows, 6) = "���ް���(\)":      .ColWidth(6) = 1700:       .ColAlignment(6) = flexAlignRightCenter:      .FixedAlignment(6) = flexAlignCenterCenter
        .TextMatrix(nRows, 7) = "�ΰ���(\)":        .ColWidth(7) = 1500:       .ColAlignment(7) = flexAlignRightCenter:      .FixedAlignment(7) = flexAlignCenterCenter
        .TextMatrix(nRows, 8) = "û���ݾ�(\)":      .ColWidth(8) = 1800:       .ColAlignment(8) = flexAlignRightCenter:      .FixedAlignment(8) = flexAlignCenterCenter
        
        .TextMatrix(nRows, 9) = "CustomID":         .ColWidth(9) = 0:          .ColAlignment(9) = flexAlignCenterCenter:     .FixedAlignment(9) = flexAlignCenterCenter
        .TextMatrix(nRows, 10) = "Depth":           .ColWidth(10) = 0:         .ColAlignment(10) = flexAlignCenterCenter:    .FixedAlignment(10) = flexAlignCenterCenter
        
        .MergeCells = flexMergeFree
        
        .ExtendLastCol = False
        
        .Redraw = flexRDDirect

        .MergeCells = flexMergeFree
        For i = 0 To .FixedRows - 1
            .MergeRow(i) = True
        Next i
        Call SetPrintMode(grdData(1), 1, True)
    End With
    
End Sub

Sub SetDataToPrn(ByVal kCustom As String)
    Dim II%, JJ%
    
    Call FillGrdPrintHeader(kCustom)
    With grdData(1)
        For II = grdData(0).FixedRows To grdData(0).Rows - 1
            If Trim(grdData(0).TextMatrix(II, 1)) = Trim(kCustom) Then
                .AddItem ""
                .RowHeight(.Rows - 1) = 400
                For JJ = 2 To .Cols - 1
                    .TextMatrix(.Rows - 1, JJ) = grdData(0).TextMatrix(II, JJ)
                Next JJ
            End If
            .Redraw = flexRDDirect
            .RowHidden(.Rows - 1) = grdData(0).RowHidden(II)
        Next II
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = " "
        .RowHeight(.Rows - 1) = 200
        .MergeRow(.Rows - 1) = True
        
        ' ��������, �ΰ��� font = 8�� �缳��
'        For II = .FixedRows To .Rows - 1
'            .Cell(flexcpFontSize, II, 0, II, 4) = 8
'            .Cell(flexcpFontSize, II, 0, II, 11) = 8
'        Next II
        
        Call SetPrintMode(grdData(1), 1, True)
        
'        .AddItem ""
'        .RowHeight(.Rows - 1) = 500
        
'        .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 5) = "���ް���: " & SetCurrency(.ValueMatrix(.Rows - 3, 10), 0) & "  ��"
'
'        .Cell(flexcpText, .Rows - 1, 6, .Rows - 1, 9) = "�ΰ���: " & SetCurrency(.ValueMatrix(.Rows - 3, 11), 0) & "  ��"
'
'        .Cell(flexcpText, .Rows - 1, 10, .Rows - 1, 11) = "�ѱݾ�: " & SetCurrency(.ValueMatrix(.Rows - 3, 10) + .ValueMatrix(.Rows - 3, 11), 0) & "  ��"
        
'        .Cell(flexcpAlignment, .Rows - 1, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
'        .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
'        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = PRNHeaderColor
                                             
        .MergeRow(.Rows - 1) = True
        
        .MergeCells = flexMergeFree
        For II = 1 To 5
            .MergeCol(II) = True
        Next II
                
        .Redraw = flexRDDirect
        grdData(1).PrintGrid "��������", True, 2, 1500, 500
        
'        Call SetPrintMode(grdData(1), 1, False)
        
    End With
End Sub


Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False

End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 13605, 9660

    Call InitGrid(0)
    Call FillGrdPrintHeader("��������")
    Call SetOperate(Me)
    
    '----- ��¥����
    dtpDate(0) = Now
    dtpDate(1) = Now
    
'    CboStuffClss2.ListIndex = 0
    
    '----- �˻��� �԰��� ����
    '----- 0:A��, 1:B��, 2:�ð��� 3.����
''    With CboOrderFlag
''        .AddItem "1.LOCAL":            .ItemData(0) = 0      ' A�� + �ð���
''        .AddItem "2.����":             .ItemData(1) = 1      ' B��
''        .ListIndex = 0
''    End With
    
    With cboTaxClss
        .AddItem "0.����"
        .AddItem "1.���"
        .AddItem "9.��ü"
        .ListIndex = 0
    End With

    
    '--- find ��Ʈ�� icon����
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    'S_201211_��������_02 �� ���� �߰�
    cmdOrderPrint.Picture = LoadResPicture("PRINT", vbResIcon)
    

    cmdFind(0).Enabled = False
    
    txtCustom(1).Enabled = False
    pnlPrint.Visible = False

End Sub

Private Sub InitGrid(ByVal Index As Integer)
    Dim II%, nRows As Integer
    
    Call SetVSFlexGrid(grdData(Index))
    With grdData(Index)
        .Cols = 11
        .Rows = 7
        .FixedRows = .Rows
        .FixedCols = 1
        
        nRows = .Rows - 1
        For II = 0 To nRows - 1
            .RowHidden(II) = True
        Next II
        
        .RowHeight(nRows) = 400
        
        .TextMatrix(nRows, 0) = "":                 .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "�ŷ�ó":           .ColWidth(1) = 2200:       .ColAlignment(1) = flexAlignLeftCenter:       .FixedAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "ǰ��":             .ColWidth(2) = 2500:       .ColAlignment(2) = flexAlignLeftCenter:       .FixedAlignment(2) = flexAlignCenterCenter
        .TextMatrix(nRows, 3) = "��������":         .ColWidth(3) = 1400:       .ColAlignment(3) = flexAlignCenterCenter:     .FixedAlignment(3) = flexAlignCenterCenter
        .TextMatrix(nRows, 4) = "��    ��":         .ColWidth(4) = 1200:       .ColAlignment(4) = flexAlignRightCenter:      .FixedAlignment(4) = flexAlignCenterCenter
        
        .TextMatrix(nRows, 5) = "�ܰ�":             .ColWidth(5) = 1000:       .ColAlignment(5) = flexAlignRightCenter:      .FixedAlignment(5) = flexAlignCenterCenter
        .TextMatrix(nRows, 6) = "���ް���(\)":      .ColWidth(6) = 1500:       .ColAlignment(6) = flexAlignRightCenter:      .FixedAlignment(6) = flexAlignCenterCenter
        .TextMatrix(nRows, 7) = "�ΰ���(\)":        .ColWidth(7) = 1400:       .ColAlignment(7) = flexAlignRightCenter:      .FixedAlignment(7) = flexAlignCenterCenter
        .TextMatrix(nRows, 8) = "û���ݾ�(\)":      .ColWidth(8) = 1600:       .ColAlignment(8) = flexAlignRightCenter:      .FixedAlignment(8) = flexAlignCenterCenter
        
        .TextMatrix(nRows, 9) = "CustomID":         .ColWidth(9) = 0:          .ColAlignment(9) = flexAlignCenterCenter:     .FixedAlignment(9) = flexAlignCenterCenter
        .TextMatrix(nRows, 10) = "Depth":           .ColWidth(10) = 0:         .ColAlignment(10) = flexAlignCenterCenter:    .FixedAlignment(10) = flexAlignCenterCenter
        
        .MergeCells = flexMergeFree
        
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .Redraw = flexRDDirect
    End With


    Call SetVSFlexGrid(grdDataOrder)
    With grdDataOrder
        .Cols = 13
        .Rows = 1
        .FixedRows = .Rows
        .FixedCols = 1
        
        nRows = .Rows - 1
        
        .TextMatrix(nRows, 0) = "":                 .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "�ŷ�ó":           .ColWidth(1) = 2200:       .ColAlignment(1) = flexAlignLeftCenter:       .FixedAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "ǰ��":             .ColWidth(2) = 2500:       .ColAlignment(2) = flexAlignLeftCenter:       .FixedAlignment(2) = flexAlignCenterCenter
        .TextMatrix(nRows, 3) = "OrderNo.":         .ColWidth(3) = 1400:       .ColAlignment(3) = flexAlignCenterCenter:     .FixedAlignment(3) = flexAlignCenterCenter
        .TextMatrix(nRows, 4) = "������":         .ColWidth(4) = 1200:       .ColAlignment(4) = flexAlignRightCenter:      .FixedAlignment(4) = flexAlignCenterCenter
        .TextMatrix(nRows, 5) = "��������":         .ColWidth(5) = 1400:       .ColAlignment(5) = flexAlignCenterCenter:     .FixedAlignment(5) = flexAlignCenterCenter
        .TextMatrix(nRows, 6) = "��    ��":         .ColWidth(6) = 1200:       .ColAlignment(6) = flexAlignRightCenter:      .FixedAlignment(6) = flexAlignCenterCenter
        .TextMatrix(nRows, 7) = "�ܰ�":             .ColWidth(7) = 1000:       .ColAlignment(7) = flexAlignRightCenter:      .FixedAlignment(7) = flexAlignCenterCenter
        .TextMatrix(nRows, 8) = "���ް���(\)":      .ColWidth(8) = 1500:       .ColAlignment(8) = flexAlignRightCenter:      .FixedAlignment(8) = flexAlignCenterCenter
        .TextMatrix(nRows, 9) = "�ΰ���(\)":        .ColWidth(9) = 1400:       .ColAlignment(9) = flexAlignRightCenter:      .FixedAlignment(9) = flexAlignCenterCenter
        .TextMatrix(nRows, 10) = "û���ݾ�(\)":     .ColWidth(10) = 1600:      .ColAlignment(10) = flexAlignRightCenter:      .FixedAlignment(10) = flexAlignCenterCenter
        .TextMatrix(nRows, 11) = "CustomID":        .ColWidth(11) = 0:         .ColAlignment(11) = flexAlignCenterCenter:     .FixedAlignment(11) = flexAlignCenterCenter
        .TextMatrix(nRows, 12) = "Depth":           .ColWidth(12) = 0:         .ColAlignment(12) = flexAlignCenterCenter:    .FixedAlignment(12) = flexAlignCenterCenter
        
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(1) = True
        .MergeCol(2) = True
        
        
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExNone
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Sub FillgrdData()
    Dim oStuffIn As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%
    Dim dDate_str As String
    Dim sDate As String, eDate As String
    Dim StuffClss As String

    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    
    Set rs = oStuffIn.GetProcCostReport(sDate, IIf(chkSearch(1) = vbChecked, 1, 0), txtCustom(1).Tag, Left(cboTaxClss, 1))

    Set oStuffIn = Nothing
    cboCustom.Clear
    
    
    ''S_201203_��������_02 �� ���� ����
''    cboCustom.AddItem AllStr
    cboCustom.AddItem AllStr & Space(50) & CheckNull(rs!CustomID)               '��ü �߰�-�ڵ�� 0000
    
    With grdData(0)
        .Rows = .FixedRows
        .Redraw = flexRDDirect

        If rs.RecordCount = 1 Then
            rs.Close
            Set rs = Nothing
            Screen.MousePointer = vbDefault
            MsgBox LoadResString(203), vbInformation
            Exit Sub
        Else
            Do Until rs.EOF
                    
                Select Case rs!Depth
                    Case "Z1"
                        If Trim(.TextMatrix(.Rows - 1, 2)) <> Trim(rs!Article) Then
                            .AddItem "" & vbTab & Trim(rs!kCustom)
                            .RowHidden(.Rows - 1) = True
                        ElseIf Trim(.TextMatrix(.Rows - 1, 3)) <> Trim(rs!WorkName) Then
                            .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article)
                            .RowHidden(.Rows - 1) = True
                        End If
                    Case "Z2"
                        .AddItem "" & vbTab & Trim(rs!kCustom)
                        .RowHidden(.Rows - 1) = True
                        
                        'S_201203_��������_02 �� ���� ����
'                        cboCustom.AddItem Trim(rs!kCustom)
                        cboCustom.AddItem CheckNull(Trim(rs!kCustom)) & Space(50) & CheckNull(rs!CustomID)

                    Case "Z3"
                        .AddItem ""
                        .RowHidden(.Rows - 1) = True
                End Select
                
                .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & vbTab & rs!WorkName & vbTab & _
                        SetCurrency(rs!SumQtyYDS, 0) & vbTab & SetCurrency(rs!UnitPrice, 0) & vbTab & _
                        SetCurrency(rs!AmountWon, 0) & vbTab & SetCurrency(rs!TaxPrice, 0) & vbTab & SetCurrency(rs!TotalPrice, 0) & vbTab & _
                        rs!CustomID & vbTab & rs!Depth
                        
                Select Case rs!Depth
                    Case "Z1"
                        If rs!OrderFlag = "0" And rs!TaxClss = "������" And rs!TaxPrice = 0 Then
                            Select Case rs!DealClss
                                Case "0":  .TextMatrix(.Rows - 1, 7) = ""
                                Case "1":  .TextMatrix(.Rows - 1, 7) = "LC/OPEN"
                                Case "2":  .TextMatrix(.Rows - 1, 7) = "���Ž��μ�"
                                Case "3":  .TextMatrix(.Rows - 1, 7) = "�Ӱ�����༭"
                            End Select
                        End If
                    Case "Z2"
                        .TextMatrix(.Rows - 1, 2) = "   ��       ��   "
                        .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = PRNHeaderColor
                        .Cell(flexcpFontBold, .Rows - 1, 2, .Rows - 1, .Cols - 1) = True
                
                    Case "Z3"
                        .TextMatrix(.Rows - 1, 1) = "   ��       ��   "
                        .TextMatrix(.Rows - 1, 2) = ""
                        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = PRNHeaderColor
                        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                End Select
                
                rs.MoveNext
            Loop
        End If
        
'        If txtCustom(1).Tag <> "" Then
'            .RowHidden(.Rows - 1) = True
'        End If
        
        .MergeCells = flexMergeFree
        For i = 1 To 5
            .MergeCol(i) = True
        Next i
        
        .Redraw = flexRDDirect
    End With
    cboCustom.ListIndex = 0         '��ü ����
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "FrmProcCostReport.FillGrdData", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub


Private Sub optPrn_Click(Index As Integer)
    If Index = 0 Then
        cboCustom.Enabled = False
    Else
        cboCustom.Enabled = True
    End If
End Sub



Private Sub txtCustom_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 1
            Call MoveFocus(KeyCode)
    End Select

End Sub

Private Sub txtCustom_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 1
            If KeyAscii = vbKeyReturn Then
                Call cmdFind_Click(0)
            End If
    End Select
End Sub


'S_201203_��������_02 �� ���� �߰�
'1���� ���þ�ü ��¸� ��
Private Sub ExcelPrintByOneCustByOrder(Optional bPreview As Boolean = True)

    Dim oExcel                          As Excel.Application
    Dim oExcelBook                      As Excel.Workbook
    Dim oExcelSheet                     As Excel.Worksheet
    Dim oFs         As New FileSystemObject
    Dim sReport As String
    Dim sCustomID As String
    Dim sCustom As String
    Dim sArticleID As String
    Dim sArticle As String
    Dim oStuffIn As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%
    Dim dDate_str As String
    Dim sDate As String, eDate As String
    Dim StuffClss As String

    Dim nRow%, nCol%, nPage%, nBaseRow%
    Dim iExcelByPage As Integer                 '�� ���̴� ��µǴ� ���� ROW�� : 32��

    Dim nChkCustom As Integer

    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName

    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    
    'S_201203_��������_02 �� ���� ���� -OLD
''    nChkCustom = 1
''    sCustomID = Trim(Mid(cboCustom.List(cboCustom.ListIndex), 50))                     '�ŷ�ó�ڵ�
''    sCustom = RTrim(Left(cboCustom.List(cboCustom.ListIndex), 50))                     '�ŷ�ó��
''
''
    '�ŷ�ó ����
    'S_201203_��������_02 �� ���� ���� -NEW
    If (optGub(0).Value) = True Or (chkSearch(1).Value = 1 And txtCustom(1).Tag <> "" And txtCustom(1) <> "") Then
        nChkCustom = 1
        sCustomID = IIf(txtCustom(1).Tag <> "", txtCustom(1).Tag, grdData(0).TextMatrix(grdData(0).Row, 9))      '�ŷ�ó�ڵ�
        sCustom = IIf(txtCustom(1).Tag <> "", txtCustom(1).Text, grdData(0).TextMatrix(grdData(0).Row, 1))           '�ŷ�ó��
    End If
    
    Set rs = oStuffIn.GetProcCostReportOrder(sDate, nChkCustom, sCustomID, Left(cboTaxClss, 1))

    '-------------------------------------------------------------------------------
    '������ Ȯ�ο� �׸��� ä���
    With grdDataOrder
        .Rows = .FixedRows
        .Redraw = flexRDDirect

        If rs.RecordCount = 1 Then
            rs.Close
            Set rs = Nothing
            Screen.MousePointer = vbDefault
            MsgBox LoadResString(203), vbInformation
            Exit Sub
        Else
            Do Until rs.EOF

'''                Select Case rs!Depth
'''                    Case "Z1"
'''                        If Trim(.TextMatrix(.Rows - 1, 2)) <> Trim(rs!Article) Then
'''                            .AddItem "" & vbTab & Trim(rs!kCustom)
'''                            .RowHidden(.Rows - 1) = True
'''                        ElseIf Trim(.TextMatrix(.Rows - 1, 3)) <> Trim(rs!WorkName) Then
'''                            .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article)
'''                            .RowHidden(.Rows - 1) = True
'''                        End If
'''                    Case "Z2"
'''                        .AddItem "" & vbTab & Trim(rs!kCustom)
'''                        .RowHidden(.Rows - 1) = True
'''
'''                    Case "Z3"
'''                        .AddItem ""
'''                        .RowHidden(.Rows - 1) = True
'''                End Select

                'S_201212_��������_06 �� ���� ���� - ���ַ� ���� ǥ�ø� ���� unitClss ���� �߰�
                .AddItem "" & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & vbTab & Trim(rs!OrderNo) & vbTab & _
                        IIf(CheckNum(rs!OrderQty) = 0, "", Format(CheckNum((rs!OrderQty)), "#,##0") & IIf(rs!UnitClss = "1", " M", " Y")) & vbTab & rs!WorkName & vbTab & _
                        SetCurrency(rs!SumQtyYDS, 0) & vbTab & SetCurrency(rs!UnitPrice, 0) & vbTab & _
                        SetCurrency(rs!AmountWon, 0) & vbTab & SetCurrency(rs!TaxPrice, 0) & vbTab & SetCurrency(rs!TotalPrice, 0) & vbTab & _
                        rs!CustomID & vbTab & rs!Depth

                Select Case rs!Depth
                    Case "Z1"
                        If rs!OrderFlag = "0" And rs!TaxClss = "������" And rs!TaxPrice = 0 Then
                            Select Case rs!DealClss
                                Case "0":  .TextMatrix(.Rows - 1, 7) = ""
                                Case "1":  .TextMatrix(.Rows - 1, 7) = "LC/OPEN"
                                Case "2":  .TextMatrix(.Rows - 1, 7) = "���Ž��μ�"
                                Case "3":  .TextMatrix(.Rows - 1, 7) = "�Ӱ�����༭"
                            End Select
                        End If
                    Case "Z2"
                        .TextMatrix(.Rows - 1, 2) = "   ��       ��   "
                        .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = PRNHeaderColor
                        .Cell(flexcpFontBold, .Rows - 1, 2, .Rows - 1, .Cols - 1) = True

                    Case "Z3"
                        .TextMatrix(.Rows - 1, 1) = "   ��       ��   "
                        .TextMatrix(.Rows - 1, 2) = ""
                        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = PRNHeaderColor
                        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                End Select

                rs.MoveNext
            Loop

            rs.MoveFirst        '���ڵ�� ���� ��ġ�� �̵�
        End If

        .MergeCells = flexMergeFree
        For i = 1 To 5
            .MergeCol(i) = True
        Next i

        .Redraw = flexRDDirect
    End With
    '-------------------------------------------------------------------------------


    Set oStuffIn = Nothing

    iExcelByPage = 32                 '������ ���������� �� ����

    Set oExcel = New Excel.Application
    Set oExcelBook = oExcel.Workbooks.Open(App.Path & Reportfile_Excel_Order)

    '������ �Ʒ� �ּ� ����---------------------------------------------------
''    oExcel.WindowState = xlMaximized
''    oExcel.Application.Visible = True
    '--------------------------------------------------------------------------

    Set oFs = New FileSystemObject
    
    

    'û���� ���� ���� ��� ����
    If Not oFs.FolderExists(CStr(App.Path) & "\û����\") Then
        oFs.CreateFolder (CStr(App.Path) & "\û����\")           '������� ���� ����
    End If
    

    'S_201203_��������_02 �� ���� ����(OLD:_ProcCost.xls)
    sReport = App.Path & "\û����\������û����_" & sDate & "_" & sCustom & ".xls"
    
    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)

    nPage = 1
    nBaseRow = 0
    nRow = 0

    '-----------------------------------------------------------------------------------------------------------------------------------------------------
    ' ���� �ڷ� ��� �κ� Ȯ�� �� ����
    '34�� ���� ���������� ��µǴ� ������ ��
    With oExcel

        ' �ű� ������ ����
        .Worksheets("Form").Activate

        '****Excel PageHeader Start========================================================================================
''        If Trim(txtCustom(1).Tag) <> "" Then
''            .Cells(4, 5) = Trim(txtCustom(1).Text)     '�ŷ�ó
''        Else
''            .Cells(4, 5) = grdData(0).TextMatrix(grdData(0).Row, 1)     '�ŷ�ó
''        End If

        '*********************************************************************************
        '** ������
        '---------------------------------------------------------------------------------
        .Cells(4, 7) = sCustom                                                      '�ŷ�ó
        .Cells(5, 7) = Left(sDate, 4) & "�� " & Right(sDate, 2) & "��"              '��������
        
''        'S_201212_��������_06 �� ���� ����-�ּ�ó��
''        .Cells(5, 30) = MakeDate(DF_LONG, dtpDate(1))                                '�������
        .Cells(7, 7) = Format(g_companyInfo.Company_No, "###-##-#####")             '��Ϲ�ȣ
        .Cells(7, 25) = g_companyInfo.Company_Name                                  '��ȣ
        .Cells(7, 39) = g_companyInfo.Chief                                         '��ǥ��

''        'S_201312_��������_99 �� ���� ����-OLD�ҽ�
''        .Cells(8, 7) = g_companyInfo.Address1 & " " & g_companyInfo.Address2        '������ּ�
        'S_201312_��������_99 �� ���� ����-NEW�ҽ�
        If CheckNull(g_companyInfo.Address1) <> "" Then            '���θ� �ּҰ� �������
            .Cells(8, 7) = g_companyInfo.Address1 & " " & g_companyInfo.Address2        '������ּ�
        Else
            .Cells(8, 7) = g_companyInfo.AddressJiBun1 & " " & g_companyInfo.AddressJiBun2        '������ּ�
        End If
        
        .Cells(8, 25) = g_companyInfo.Company_type                                  '����
        .Cells(8, 39) = g_companyInfo.Category                                      '����
        '*********************************************************************************
        
        '****Excel PageHeader End========================================================================================
        nBaseRow = GetExcelBaseRowByOrder(nPage)
        nRow = 0

        '������ �߰�
        Call InsertExcelFormByOrder(oExcel, nPage, 1)
        .Worksheets("Report").Activate

        Do Until rs.EOF

            If nRow >= iExcelByPage Then             'nRow�� 0���� �����ϹǷ� 32>=32  �� ��� ������ ����
               nPage = nPage + 1
               Call InsertExcelFormByOrder(oExcel, nPage, 1)
               nBaseRow = GetExcelBaseRowByOrder(nPage)
               nRow = 0
               sArticle = ""
            End If

            If rs!Depth <> "Z2" Then            '1����ü �����̸鼭 �ŷ�ó�谡 �ƴ� ��츸 ���

                If rs!Depth = "Z3" Then         '�Ѱ�

                    '�ϴ� �հ�κ� Merge
                    Call ExcelTotalByOrder(oExcel, nPage, nBaseRow, nRow, 1)

                    .Cells(10 + nBaseRow + nRow, 2) = "�Ѱ�"          ' �Ѱ�
                    .Cells(10 + nBaseRow + nRow, 11) = ""  'OrderNo
                    .Cells(10 + nBaseRow + nRow, 30) = ""  '�ܰ�

                Else

                    If sArticle <> Trim(rs!Article) Then
                        .Cells(10 + nBaseRow + nRow, 2) = Trim(rs!Article)   'ǰ��
                        sArticle = Trim(rs!Article)
                    Else
                        .Cells(10 + nBaseRow + nRow, 2) = ""                    'ǰ��
                    End If

''                    .Cells(10 + nBaseRow + nRow, 23) = rs!WorkName   '��������

                    .Cells(10 + nBaseRow + nRow, 11) = CheckNull(rs!OrderNo)   'OrderNo
                    .Cells(10 + nBaseRow + nRow, 30) = SetCurrency(rs!UnitPrice, 0)  '�ܰ�
                End If


                'S_201212_��������_06 �� ���� ����-OLD �ҽ�
''                .Cells(10 + nBaseRow + nRow, 19) = Format(CheckNum(rs!OrderQty), "#,##0")  '���ַ�
                
                'S_201212_��������_06 �� ���� ����-NEW �ҽ�
                .Cells(10 + nBaseRow + nRow, 19) = IIf(CheckNum(rs!OrderQty) = 0, "", Format(CheckNum((rs!OrderQty)), "#,##0") & IIf(rs!UnitClss = "1", " M", " Y"))

                .Cells(10 + nBaseRow + nRow, 23) = rs!WorkName   '��������
                .Cells(10 + nBaseRow + nRow, 27) = SetCurrency(rs!SumQtyYDS, 0)   '����
''                .Cells(10 + nBaseRow + nRow, 30) = SetCurrency(rs!UnitPrice, 0)  '�ܰ�
                .Cells(10 + nBaseRow + nRow, 33) = SetCurrency(rs!AmountWon, 0)   '���ް���
                .Cells(10 + nBaseRow + nRow, 38) = SetCurrency(rs!TaxPrice, 0)   '�ΰ���
                .Cells(10 + nBaseRow + nRow, 42) = SetCurrency(rs!TotalPrice, 0)  'û���ݾ�

                nRow = nRow + 1
            End If
            '---------------------------------------------------------------------------

            rs.MoveNext
        Loop
    End With

    Set oFs = Nothing

    Call oExcelBook.SaveAs(sReport)
    
    
    If bPreview Then                    '�̸����� ���
        oExcel.WindowState = xlMaximized
        oExcel.Application.Visible = True
        oExcel.ActiveWindow.SelectedSheets.PrintPreview
    Else                                '�ٷ��μ�
        oExcel.ActiveWindow.SelectedSheets.PrintOut Copies:=1
        Call ProcessClose("XLMAIN")
    End If


    Screen.MousePointer = vbDefault

    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    Set oFs = Nothing

    Exit Sub

ErrHandler:

    Call ErrorBox(Err.Number, "FrmProcCostReport.ExcelPrintByOneCustByOrder", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing

    If Err.Number <> 0 Then
        MsgBox Err.Number & " : " & Err.Description
    End If


End Sub

'S_201203_��������_02 �� ���� �߰�
'�հ� �κ� Merge
Private Function ExcelTotalByOrder(oExcel As Excel.Application, nPage As Integer, nBaseRow As Integer, nRow As Integer, nPrnGub As Integer)

    On Error GoTo Err_Rtn

    With oExcel

        If nPrnGub = 1 Then     'Ư���ŷ�ó ������ ���

            '�Ѱ�
           .Range("B" & 10 + nBaseRow + nRow & ":Z" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

''            '��������
''           .Range("G" & 10 + nBaseRow + nRow & ":I" & 10 + nBaseRow + 31).Select
''           With .Selection
''               .HorizontalAlignment = xlCenter
''               .VerticalAlignment = xlTop
''               .WrapText = True
''               .Orientation = 0
''               .AddIndent = False
''               .ShrinkToFit = False
''               .Borders(xlEdgeRight).LineStyle = xlContinuous
''               .Borders(xlEdgeRight).Weight = xlHairline
''               .Font.Size = 10
''               .WrapText = False
''               .ShrinkToFit = True
''               .Font.Bold = True
''           End With
''           .Selection.Merge

           '����
           .Range("AA" & 10 + nBaseRow + nRow & ":AC" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '�ܰ�
           .Range("AD" & 10 + nBaseRow + nRow & ":AF" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '���ް���
           .Range("AG" & 10 + nBaseRow + nRow & ":AK" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '�ΰ�����
           .Range("AL" & 10 + nBaseRow + nRow & ":AO" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '�ѱݾ�
           .Range("AP" & 10 + nBaseRow + nRow & ":AT" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           .Selection.Interior.ColorIndex = xlNone



        Else                            '������ü ���-1�� ��ü�� �����
''            '�Ѱ�
''           .Range("B" & 8 + nBaseRow + nRow & ":I" & 8 + nBaseRow + 33).Select
''           With .Selection
''               .HorizontalAlignment = xlCenter
''               .VerticalAlignment = xlTop
''               .WrapText = True
''               .Orientation = 0
''               .AddIndent = False
''               .ShrinkToFit = False
''               .Borders(xlEdgeRight).LineStyle = xlContinuous
''               .Borders(xlEdgeRight).Weight = xlHairline
''               .Font.Size = 10
''               .WrapText = False
''               .ShrinkToFit = True
''               .Font.Bold = True
''           End With
''           .Selection.Merge
''
''            '��������
''           .Range("J" & 8 + nBaseRow + nRow & ":K" & 8 + nBaseRow + 33).Select
''           With .Selection
''               .HorizontalAlignment = xlCenter
''               .VerticalAlignment = xlTop
''               .WrapText = True
''               .Orientation = 0
''               .AddIndent = False
''               .ShrinkToFit = False
''               .Borders(xlEdgeRight).LineStyle = xlContinuous
''               .Borders(xlEdgeRight).Weight = xlHairline
''               .Font.Size = 10
''               .WrapText = False
''               .ShrinkToFit = True
''               .Font.Bold = True
''           End With
''           .Selection.Merge
''
''           '����
''           .Range("L" & 8 + nBaseRow + nRow & ":M" & 8 + nBaseRow + 33).Select
''           With .Selection
''               .HorizontalAlignment = xlRight
''               .VerticalAlignment = xlTop
''               .WrapText = True
''               .Orientation = 0
''               .AddIndent = False
''               .ShrinkToFit = False
''               .Borders(xlEdgeRight).LineStyle = xlContinuous
''               .Borders(xlEdgeRight).Weight = xlHairline
''               .Font.Size = 10
''               .WrapText = False
''               .ShrinkToFit = True
''               .Font.Bold = True
''           End With
''           .Selection.Merge
''
''           '�ܰ�
''           .Range("N" & 8 + nBaseRow + nRow & ":O" & 8 + nBaseRow + 33).Select
''           With .Selection
''               .HorizontalAlignment = xlRight
''               .VerticalAlignment = xlTop
''               .WrapText = True
''               .Orientation = 0
''               .AddIndent = False
''               .ShrinkToFit = False
''               .Borders(xlEdgeRight).LineStyle = xlContinuous
''               .Borders(xlEdgeRight).Weight = xlHairline
''               .Font.Size = 10
''               .WrapText = False
''               .ShrinkToFit = True
''               .Font.Bold = True
''           End With
''           .Selection.Merge
''
''           '���ް���
''           .Range("P" & 8 + nBaseRow + nRow & ":R" & 8 + nBaseRow + 33).Select
''           With .Selection
''               .HorizontalAlignment = xlRight
''               .VerticalAlignment = xlTop
''               .WrapText = True
''               .Orientation = 0
''               .AddIndent = False
''               .ShrinkToFit = False
''               .Borders(xlEdgeRight).LineStyle = xlContinuous
''               .Borders(xlEdgeRight).Weight = xlHairline
''               .Font.Size = 10
''               .WrapText = False
''               .ShrinkToFit = True
''               .Font.Bold = True
''           End With
''           .Selection.Merge
''
''           '�ΰ�����
''           .Range("S" & 8 + nBaseRow + nRow & ":U" & 8 + nBaseRow + 33).Select
''           With .Selection
''               .HorizontalAlignment = xlRight
''               .VerticalAlignment = xlTop
''               .WrapText = True
''               .Orientation = 0
''               .AddIndent = False
''               .ShrinkToFit = False
''               .Borders(xlEdgeRight).LineStyle = xlContinuous
''               .Borders(xlEdgeRight).Weight = xlHairline
''               .Font.Size = 10
''               .WrapText = False
''               .ShrinkToFit = True
''               .Font.Bold = True
''           End With
''           .Selection.Merge
''
''           '�ѱݾ�
''           .Range("V" & 8 + nBaseRow + nRow & ":X" & 8 + nBaseRow + 33).Select
''           With .Selection
''               .HorizontalAlignment = xlRight
''               .VerticalAlignment = xlTop
''               .WrapText = True
''               .Orientation = 0
''               .AddIndent = False
''               .ShrinkToFit = False
''               .Borders(xlEdgeRight).LineStyle = xlContinuous
''               .Font.Size = 10
''               .WrapText = False
''               .ShrinkToFit = True
''               .Font.Bold = True
''           End With
''           .Selection.Merge
''
''           .Selection.Interior.ColorIndex = xlNone


        End If
    End With

    Exit Function

Err_Rtn:
    If Err.Number <> 0 Then MsgBox Err.Number & "," & Err.Description, vbCritical, "[ExcelTotal]"
End Function

'S_201203_��������_02 �� ���� �߰�
'�ŷ�ó �հ�-��ü ����� ��츸 ����
Private Function ExcelSubTotalByOrder(oExcel As Excel.Application, nPage As Integer, nBaseRow As Integer, nRow As Integer, nPrnGub As Integer)

    On Error GoTo Err_Rtn

    With oExcel

        If nPrnGub = 0 Then
           '�ŷ�ó��
           .Range("B" & 8 + nBaseRow + nRow & ":I" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '��������
           .Range("J" & 8 + nBaseRow + nRow & ":K" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '����
           .Range("L" & 8 + nBaseRow + nRow & ":M" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '�ܰ�
           .Range("N" & 8 + nBaseRow + nRow & ":O" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '���ް���
           .Range("P" & 8 + nBaseRow + nRow & ":R" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '�ΰ�����
           .Range("S" & 8 + nBaseRow + nRow & ":U" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           '�ѱݾ�
           .Range("V" & 8 + nBaseRow + nRow & ":X" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge

           .Selection.Interior.ColorIndex = xlNone

        End If

    End With

    Exit Function

Err_Rtn:
    If Err.Number <> 0 Then MsgBox Err.Number & "," & Err.Description, vbCritical, "[ExcelsubTotal]"
End Function

'S_201203_��������_02 �� ���� �߰�
'BaseRow��ȯ �Լ�
Private Function GetExcelBaseRowByOrder(nPage)
    GetExcelBaseRowByOrder = (nPage - 1) * EXCEL_ROW
End Function

'S_201203_��������_02 �� ���� �߰�
'���������� ����-��Ʈ����
Private Function InsertExcelFormByOrder(oExcel As Excel.Application, nPage As Integer, nPrnGub As Integer)
    Dim i%, nBaseRow%

    nBaseRow = GetExcelBaseRowByOrder(nPage)
    With oExcel
        If nPrnGub = 1 Then     'Ư���ŷ�ó ������ ���
            .Sheets("Form").Select

        Else
            .Sheets("Form2").Select         '�����ŷ�ó
        End If

        .Rows("1:" & CStr(EXCEL_ROW)).Select
        .Selection.Copy

        .Sheets("Report").Select
        .Rows(CStr(nBaseRow + 1) & ":" & CStr(nBaseRow + 1)).Select
        .Selection.Insert Shift:=xlDown
        
        
        'S_201212_��������_06 �� ���� �߰�-���� ������ ǥ��
        .Cells(nBaseRow + 42, 38) = "PAGE : " & nPage
    End With
End Function

'S_201211_��������_02 �� ���� �߰�
'1���� ���þ�ü ���
Private Sub ExcelPrintByOneCust(Optional bPreview As Boolean = True)

    Dim oExcel                          As Excel.Application
    Dim oExcelBook                      As Excel.Workbook
    Dim oExcelSheet                     As Excel.Worksheet
    Dim oFs         As New FileSystemObject
    Dim sReport As String
    Dim sCustomID As String
    Dim sCustom As String
    Dim sArticleID As String
    Dim sArticle As String
    Dim oStuffIn As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%
    Dim dDate_str As String
    Dim sDate As String, eDate As String
    Dim StuffClss As String

    Dim nRow%, nCol%, nPage%, nBaseRow%
    Dim iExcelByPage As Integer                 '�� ���̴� ��µǴ� ���� ROW�� : 32��
    
    Dim nChkCustom As Integer
    
    On Error GoTo ErrHandler

    
    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    
    '�ŷ�ó ����
    If (optGub(0).Value) = True Or (chkSearch(1).Value = 1 And txtCustom(1).Tag <> "" And txtCustom(1) <> "") Then
        nChkCustom = 1
        sCustomID = IIf(txtCustom(1).Tag <> "", txtCustom(1).Tag, grdData(0).TextMatrix(grdData(0).Row, 9))      '�ŷ�ó�ڵ�
        sCustom = IIf(txtCustom(1).Tag <> "", txtCustom(1).Text, grdData(0).TextMatrix(grdData(0).Row, 1))           '�ŷ�ó��
    End If
                    
''    Set rs = oStuffIn.GetProcCostReport(sDate, IIf(chkSearch(1) = vbChecked, 1, 0), txtCustom(1).Tag, Left(cboTaxClss, 1))
    Set rs = oStuffIn.GetProcCostReport(sDate, nChkCustom, sCustomID, Left(cboTaxClss, 1))
    Set oStuffIn = Nothing
    
    iExcelByPage = 32                 '�ҷ��� ���������� �� ����
    
    Set oExcel = New Excel.Application
    Set oExcelBook = oExcel.Workbooks.Open(App.Path & Reportfile_Excel)
    
    '������ �Ʒ� �ּ� ����---------------------------------------------------
''    oExcel.WindowState = xlMaximized
''    oExcel.Application.Visible = True
    '--------------------------------------------------------------------------
    
    Set oFs = New FileSystemObject
    
    'û���� ���� ���� ��� ����
    If Not oFs.FolderExists(CStr(App.Path) & "\û����\") Then
        oFs.CreateFolder (CStr(App.Path) & "\û����\")           '������� ���� ����
    End If

    
''    If Trim(txtCustom(1).Tag) <> "" Then
''        sCustomID = Trim(txtCustom(1).Text)     '�ŷ�ó
''    Else
''        sCustomID = grdData(0).TextMatrix(grdData(0).Row, 1)    '�ŷ�ó
''    End If

''    sReport = App.Path & "\Excel\" & sDate & "_" & sCustomID & "_ProcCost.xls"

    If nChkCustom = 1 Then          '1�� �ŷ�ó
    
        sReport = App.Path & "\û����\û����_" & sDate & "_" & sCustom & ".xls"
    Else
    
        sReport = App.Path & "\û����\û����_" & sDate & "_(��ü).xls"
    End If

        
    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
        
    nPage = 1
    nBaseRow = 0
    nRow = 0
        
    '-----------------------------------------------------------------------------------------------------------------------------------------------------
    ' ���� �ڷ� ��� �κ� Ȯ�� �� ����
    '34�� ���� ���������� ��µǴ� ������ ��
    With oExcel
        
        ' �ű� ������ ����
        .Worksheets("Form").Activate
        
        '****Excel PageHeader Start========================================================================================
        
        '*********************************************************************************************
        '** ������
        '---------------------------------------------------------------------------------------------
        .Cells(4, 5) = sCustom                                                      '�ŷ�ó
        .Cells(5, 5) = Left(sDate, 4) & "�� " & Right(sDate, 2) & "��"              '��������
        
        'S_201212_��������_06 �� ���� ����-�ּ�ó��
''        .Cells(5, 30) = MakeDate(DF_LONG, dtpDate(1))                                '�������
        .Cells(7, 5) = Format(g_companyInfo.Company_No, "###-##-#####")             '��Ϲ�ȣ
        .Cells(7, 15) = g_companyInfo.Company_Name                                  '��ȣ
        .Cells(7, 21) = g_companyInfo.Chief                                         '��ǥ��
        
''        'S_201312_��������_99 �� ���� ����-OLD�ҽ�
''        .Cells(8, 5) = g_companyInfo.Address1 & " " & g_companyInfo.Address2        '������ּ�
        'S_201312_��������_99 �� ���� ����-NEW�ҽ�
        If CheckNull(g_companyInfo.Address1) <> "" Then            '���θ� �ּҰ� �������
            .Cells(8, 5) = g_companyInfo.Address1 & " " & g_companyInfo.Address2        '������ּ�
        Else
            .Cells(8, 5) = g_companyInfo.AddressJiBun1 & " " & g_companyInfo.AddressJiBun2        '������ּ�
        End If

        .Cells(8, 15) = g_companyInfo.Company_type                                  '����
        .Cells(8, 21) = g_companyInfo.Category                                      '����
        .Cells(42, 4) = g_companyInfo.BANK1                                         '���¹�ȣ 20221123
        
        '*********************************************************************************************


        '****Excel PageHeader End========================================================================================
        nBaseRow = GetExcelBaseRow(nPage)
        nRow = 0
           
        '������ �߰�
        Call InsertExcelForm(oExcel, nPage, 1)
        .Worksheets("Report").Activate
            
        Do Until rs.EOF
            
            If nRow >= iExcelByPage Then             'nRow�� 0���� �����ϹǷ� 32>=32  �� ��� ������ ����
               nPage = nPage + 1
               Call InsertExcelForm(oExcel, nPage, 1)
               nBaseRow = GetExcelBaseRow(nPage)
               nRow = 0
               sArticle = ""
            End If
                    
            If rs!Depth <> "Z2" Then            '1����ü �����̸鼭 �ŷ�ó�谡 �ƴ� ��츸 ���
            
                If rs!Depth = "Z3" Then         '�Ѱ�
                    Call ExcelTotal(oExcel, nPage, nBaseRow, nRow, 1)
                    
                    .Cells(10 + nBaseRow + nRow, 2) = "�Ѱ�"          ' �Ѱ�
                    .Cells(10 + nBaseRow + nRow, 13) = ""  '�ܰ�
                Else
                
                    If sArticle <> Trim(rs!Article) Then
                        .Cells(10 + nBaseRow + nRow, 2) = Trim(rs!Article)   'ǰ��
                        sArticle = Trim(rs!Article)
                    Else
                        .Cells(10 + nBaseRow + nRow, 2) = ""                    'ǰ��
                    End If
                    
''                    .Cells(10 + nBaseRow + nRow, 7) = rs!WorkName   '��������
                    .Cells(10 + nBaseRow + nRow, 13) = SetCurrency(rs!UnitPrice, 0)  '�ܰ�
                End If
                
                .Cells(10 + nBaseRow + nRow, 7) = rs!WorkName   '��������
                .Cells(10 + nBaseRow + nRow, 10) = SetCurrency(rs!SumQtyYDS, 0)   '����
''                .Cells(10 + nBaseRow + nRow, 13) = SetCurrency(rs!UnitPrice, 0)  '�ܰ�
                .Cells(10 + nBaseRow + nRow, 15) = SetCurrency(rs!AmountWon, 0)   '���ް���
                .Cells(10 + nBaseRow + nRow, 18) = SetCurrency(rs!TaxPrice, 0)   '�ΰ���
                .Cells(10 + nBaseRow + nRow, 21) = SetCurrency(rs!TotalPrice, 0)  'û���ݾ�
                
                nRow = nRow + 1
            End If
            '---------------------------------------------------------------------------
             
            rs.MoveNext
        Loop
    End With

    Set oFs = Nothing

    Call oExcelBook.SaveAs(sReport)

    If bPreview Then                    '�̸����� ���
        oExcel.WindowState = xlMaximized
        oExcel.Application.Visible = True
        oExcel.ActiveWindow.SelectedSheets.PrintPreview
    Else                                '�ٷ��μ�
        oExcel.ActiveWindow.SelectedSheets.PrintOut Copies:=1
        Call ProcessClose("XLMAIN")
    End If
    Screen.MousePointer = vbDefault

    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    Set oFs = Nothing

    Exit Sub
    
ErrHandler:

    Call ErrorBox(Err.Number, "FrmProcCostReport.ExcelPrintByOnCust", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & " : " & Err.Description
    End If
        

End Sub

'S_201211_��������_02 �� ���� �߰�
'���������� ���þ�ü ���
Private Sub ExcelPrintByAllCust(Optional bPreview As Boolean = True)

    Dim oExcel                          As Excel.Application
    Dim oExcelBook                      As Excel.Workbook
    Dim oExcelSheet                     As Excel.Worksheet
    Dim oFs         As New FileSystemObject
    Dim sReport As String
    Dim sCustom As String
    Dim sCustomID As String
    Dim sArticleID As String
    Dim sArticle As String
    Dim oStuffIn As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%, nCheckNon%
    Dim dDate_str As String
    Dim sDate As String, eDate As String
    Dim StuffClss As String

    Dim nRow%, nCol%, nPage%, nBaseRow%
    Dim iExcelByPage As Integer                 '�� ���̴� ��µǴ� ���� ROW�� : 32��
    On Error GoTo ErrHandler

    
    Set oStuffIn = New PlusLib2.CSubul
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    sDate = Left(MakeDate(DF_SHORT, dtpDate(0)), 6)
    
    Set rs = oStuffIn.GetProcCostReport(sDate, IIf(chkSearch(1) = vbChecked, 1, 0), txtCustom(1).Tag, Left(cboTaxClss, 1))

    Set oStuffIn = Nothing
    
    
    iExcelByPage = 34                 '�ҷ��� ���������� �� ������ 35
    
    Set oExcel = New Excel.Application
    Set oExcelBook = oExcel.Workbooks.Open(App.Path & Reportfile_Excel)
    
    '������ �Ʒ� �ּ� ����---------------------------------------------------
''    oExcel.WindowState = xlMaximized
''    oExcel.Application.Visible = True
    '--------------------------------------------------------------------------

    Set oFs = New FileSystemObject
    'û���� ���� ���� ��� ����
    If Not oFs.FolderExists(CStr(App.Path) & "\û����\") Then
        oFs.CreateFolder (CStr(App.Path) & "\û����\")           '������� ���� ����
    End If
    
    '2012.12.13 ����-��ü ���� ���½ÿ��� �ŷ�ó �� �ʿ� �����Ƿ� �ּ�
''    If Trim(txtCustom(1).Tag) <> "" Then
''        sCustom = Trim(txtCustom(1).Text)     '�ŷ�ó
''    Else
''        sCustom = grdData(0).TextMatrix(grdData(0).Row, 1)    '�ŷ�ó
''    End If

''    sReport = App.Path & "\Excel\" & sDate & "_All_ProcCost.xls"
    sReport = App.Path & "\û����\û����_" & sDate & "_(��ü).xls"
        
    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
    
    nPage = 1
    nBaseRow = 0
    nRow = 0
        
    '-----------------------------------------------------------------------------------------------------------------------------------------------------
    ' ���� �ڷ� ��� �κ� Ȯ�� �� ����
    '34�� ���� ���������� ��µǴ� ������ ��
    With oExcel
        
        ' �ű� ������ ����
        .Worksheets("Form2").Activate
        
        '****Excel PageHeader Start========================================================================================
        .Cells(4, 5) = "��ü"                   '�ŷ�ó
        .Cells(5, 5) = Left(Format(dtpDate(0), "YYYYMM"), 4) & "�� " & Right(Format(dtpDate(0), "YYYYMM"), 2) & "��"  '��������
        
        '****Excel PageHeader End========================================================================================
        nBaseRow = GetExcelBaseRow(nPage)
        '������ �߰�
        Call InsertExcelForm(oExcel, nPage, 0)
        .Worksheets("Report").Activate
            
        Do Until rs.EOF

            
            If nRow >= iExcelByPage Then             'nRow�� 0���� �����ϹǷ� 32>=32 �� ��� ������ ����
               nPage = nPage + 1
               Call InsertExcelForm(oExcel, nPage, 0)
               nBaseRow = GetExcelBaseRow(nPage)
               nRow = 0
               
               sCustom = ""
               sArticle = ""
            End If
                    
            If rs!Depth = "Z1" Then            '���� �׸�
            
                If sCustom <> Trim(rs!kCustom) Then
                    .Cells(8 + nBaseRow + nRow, 2) = Trim(rs!kCustom)              '�ŷ�ó��
                    sCustom = Trim(rs!kCustom)
                Else
                    .Cells(8 + nBaseRow + nRow, 2) = ""                             '�ŷ�ó��
                End If
                
                If sArticle <> Trim(rs!Article) Then
                    .Cells(8 + nBaseRow + nRow, 6) = Trim(rs!Article)              'ǰ��
                    sArticle = Trim(rs!Article)
                Else
                    .Cells(8 + nBaseRow + nRow, 6) = ""                    'ǰ��
                End If
''                .Cells(8 + nBaseRow + nRow, 10) = rs!WorkName   '��������
                .Cells(8 + nBaseRow + nRow, 14) = SetCurrency(rs!UnitPrice, 0)  '�ܰ�
            ElseIf rs!Depth = "Z2" Then            '�ŷ�ó ��ExcelSubTotal
                Call ExcelSubTotal(oExcel, nPage, nBaseRow, nRow, 0)
                .Cells(8 + nBaseRow + nRow, 2) = "�ŷ�ó��"
                .Cells(8 + nBaseRow + nRow, 14) = ""  '�ܰ�
            
            ElseIf rs!Depth = "Z3" Then            '�Ѱ�
                Call ExcelTotal(oExcel, nPage, nBaseRow, nRow, 0)
                    
                .Cells(8 + nBaseRow + nRow, 2) = "�Ѱ�"          ' �Ѱ�
                .Cells(8 + nBaseRow + nRow, 14) = ""  '�ܰ�
            End If
  
            .Cells(8 + nBaseRow + nRow, 10) = rs!WorkName   '��������
            .Cells(8 + nBaseRow + nRow, 12) = SetCurrency(rs!SumQtyYDS, 0)   '����
''                .Cells(8 + nBaseRow + nRow, 13) = SetCurrency(rs!UnitPrice, 0)  '�ܰ�
            .Cells(8 + nBaseRow + nRow, 16) = SetCurrency(rs!AmountWon, 0)   '���ް���
            .Cells(8 + nBaseRow + nRow, 19) = SetCurrency(rs!TaxPrice, 0)   '�ΰ���
            .Cells(8 + nBaseRow + nRow, 22) = SetCurrency(rs!TotalPrice, 0)  'û���ݾ�
                
            nRow = nRow + 1
            
            '---------------------------------------------------------------------------
             
            rs.MoveNext
        Loop
    End With
    
    Set oFs = New FileSystemObject
    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
    Set oFs = Nothing

    Call oExcelBook.SaveAs(sReport)


    If bPreview Then                    '�̸����� ���
        oExcel.WindowState = xlMaximized
        oExcel.Application.Visible = True
        oExcel.ActiveWindow.SelectedSheets.PrintPreview
    Else                                '�ٷ��μ�
        oExcel.ActiveWindow.SelectedSheets.PrintOut Copies:=1
        Call ProcessClose("XLMAIN")
    End If
    
    Screen.MousePointer = vbDefault

    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    Set oFs = Nothing

    Exit Sub
    
ErrHandler:

    Call ErrorBox(Err.Number, "FrmProcCostReport.ExcelPrintByAllCust", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & " : " & Err.Description
    End If
    
    

End Sub

'S_201211_��������_02 �� ���� �߰�
'�հ� �κ� Merge
Private Function ExcelTotal(oExcel As Excel.Application, nPage As Integer, nBaseRow As Integer, nRow As Integer, nPrnGub As Integer)

    On Error GoTo Err_Rtn
    
    With oExcel
    
        If nPrnGub = 1 Then     'Ư���ŷ�ó ������ ���
        
            '�Ѱ�
           .Range("B" & 10 + nBaseRow + nRow & ":F" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
            '��������
           .Range("G" & 10 + nBaseRow + nRow & ":I" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '����
           .Range("J" & 10 + nBaseRow + nRow & ":L" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '�ܰ�
           .Range("M" & 10 + nBaseRow + nRow & ":N" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '���ް���
           .Range("O" & 10 + nBaseRow + nRow & ":Q" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '�ΰ�����
           .Range("R" & 10 + nBaseRow + nRow & ":T" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '�ѱݾ�
           .Range("U" & 10 + nBaseRow + nRow & ":X" & 10 + nBaseRow + 31).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
    
           .Selection.Interior.ColorIndex = xlNone
           
           
           
        Else                            '������ü ���
            '�Ѱ�
           .Range("B" & 8 + nBaseRow + nRow & ":I" & 8 + nBaseRow + 33).Select
           With .Selection
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
            '��������
           .Range("J" & 8 + nBaseRow + nRow & ":K" & 8 + nBaseRow + 33).Select
           With .Selection
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '����
           .Range("L" & 8 + nBaseRow + nRow & ":M" & 8 + nBaseRow + 33).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '�ܰ�
           .Range("N" & 8 + nBaseRow + nRow & ":O" & 8 + nBaseRow + 33).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '���ް���
           .Range("P" & 8 + nBaseRow + nRow & ":R" & 8 + nBaseRow + 33).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '�ΰ�����
           .Range("S" & 8 + nBaseRow + nRow & ":U" & 8 + nBaseRow + 33).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '�ѱݾ�
           .Range("V" & 8 + nBaseRow + nRow & ":X" & 8 + nBaseRow + 33).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlTop
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
    
           .Selection.Interior.ColorIndex = xlNone
            
           
        End If
    End With
    
    Exit Function
    
Err_Rtn:
    If Err.Number <> 0 Then MsgBox Err.Number & "," & Err.Description, vbCritical, "[ExcelTotal]"
End Function

'S_201211_��������_02 �� ���� �߰�
'�ŷ�ó �հ�-��ü ����� ��츸 ����
Private Function ExcelSubTotal(oExcel As Excel.Application, nPage As Integer, nBaseRow As Integer, nRow As Integer, nPrnGub As Integer)

    On Error GoTo Err_Rtn
    
    With oExcel
        
        If nPrnGub = 0 Then
           '�ŷ�ó��
           .Range("B" & 8 + nBaseRow + nRow & ":I" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '��������
           .Range("J" & 8 + nBaseRow + nRow & ":K" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '����
           .Range("L" & 8 + nBaseRow + nRow & ":M" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '�ܰ�
           .Range("N" & 8 + nBaseRow + nRow & ":O" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '���ް���
           .Range("P" & 8 + nBaseRow + nRow & ":R" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '�ΰ�����
           .Range("S" & 8 + nBaseRow + nRow & ":U" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Borders(xlEdgeRight).Weight = xlHairline
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
           
           '�ѱݾ�
           .Range("V" & 8 + nBaseRow + nRow & ":X" & 8 + nBaseRow + nRow).Select
           With .Selection
               .HorizontalAlignment = xlRight
               .VerticalAlignment = xlCenter
               .WrapText = True
               .Orientation = 0
               .AddIndent = False
               .ShrinkToFit = False
               .Borders(xlEdgeRight).LineStyle = xlContinuous
               .Font.Size = 10
               .WrapText = False
               .ShrinkToFit = True
               .Font.Bold = True
           End With
           .Selection.Merge
    
           .Selection.Interior.ColorIndex = xlNone
            
        End If
      
    End With
    
    Exit Function
    
Err_Rtn:
    If Err.Number <> 0 Then MsgBox Err.Number & "," & Err.Description, vbCritical, "[ExcelsubTotal]"
End Function

'S_201211_��������_02 �� ���� �߰�
'BaseRow��ȯ �Լ�
Private Function GetExcelBaseRow(nPage)
    GetExcelBaseRow = (nPage - 1) * EXCEL_ROW
End Function

'S_201211_��������_02 �� ���� �߰�
'���������� ����-��Ʈ����
Private Function InsertExcelForm(oExcel As Excel.Application, nPage As Integer, nPrnGub As Integer)
    Dim i%, nBaseRow%

    nBaseRow = GetExcelBaseRow(nPage)
    With oExcel
        If nPrnGub = 1 Then     'Ư���ŷ�ó ������ ���
            .Sheets("Form").Select

        Else
            .Sheets("Form2").Select         '�����ŷ�ó
        End If

        .Rows("1:" & CStr(EXCEL_ROW)).Select
        .Selection.Copy

        .Sheets("Report").Select
        .Rows(CStr(nBaseRow + 1) & ":" & CStr(nBaseRow + 1)).Select
        .Selection.Insert Shift:=xlDown
        
        'S_201212_��������_06 �� ���� �߰�-���� ������ ǥ��
        .Cells(nBaseRow + 42, 19) = "PAGE : " & nPage
    End With
End Function


