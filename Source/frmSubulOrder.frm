VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSubulOrder 
   Caption         =   "Order�� ���Ҹ���"
   ClientHeight    =   9390
   ClientLeft      =   3225
   ClientTop       =   1605
   ClientWidth     =   15225
   Icon            =   "frmSubulOrder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   15225
   Begin Threed.SSPanel pnlPrn 
      Height          =   3225
      Left            =   4680
      TabIndex        =   8
      Top             =   2850
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   5689
      _Version        =   196609
      ForeColor       =   16761024
      BackColor       =   16761024
      PictureMaskColor=   16711680
      BevelWidth      =   2
      FloodColor      =   16711935
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel SSPanel4 
         Height          =   315
         Left            =   1050
         TabIndex        =   13
         Top             =   1500
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "�μ����"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.ComboBox cboCustom 
         Height          =   300
         Left            =   2340
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   12
         Top             =   1500
         Width           =   2325
      End
      Begin Threed.SSCommand cmdPrnCancel 
         Height          =   495
         Left            =   3030
         TabIndex        =   11
         Top             =   2490
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "���"
      End
      Begin Threed.SSCommand cmdPrnOK 
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   2490
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   196609
         Caption         =   "�μ�"
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   90
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   767
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "Order�� ���Ҹ���"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   735
         Left            =   2340
         TabIndex        =   20
         Top             =   720
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1296
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optPrn 
            Caption         =   "��ü��Ȳ"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   22
            Top             =   120
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optPrn 
            Caption         =   "�����μ�"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   21
            Top             =   420
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Index           =   0
         Left            =   1050
         TabIndex        =   23
         Top             =   720
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "�μⱸ��"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   3105
         Left            =   60
         Top             =   60
         Width           =   5745
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7575
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   1050
      Width           =   15195
      _cx             =   26802
      _cy             =   13361
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
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1826
      _Version        =   196609
      Begin VB.CheckBox chkStockHidden 
         Caption         =   "�̿��ڷ� ����"
         Height          =   255
         Left            =   4470
         TabIndex        =   34
         Top             =   720
         Value           =   1  'Ȯ��
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkKG 
         Caption         =   "KG��� ��ȸ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6360
         TabIndex        =   33
         Top             =   750
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   0
         Left            =   6000
         TabIndex        =   25
         Top             =   60
         Width           =   2235
      End
      Begin VB.ComboBox CboOrderFlag 
         Height          =   300
         Left            =   6000
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   18
         Top             =   390
         Width           =   1485
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   1590
         TabIndex        =   14
         Top             =   690
         Width           =   2235
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Index           =   1
         Left            =   1590
         TabIndex        =   1
         Top             =   375
         Width           =   2235
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   90
         TabIndex        =   2
         Top             =   375
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "�� �� ó"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   3
            Top             =   45
            Width           =   1095
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   375
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
         Left            =   90
         TabIndex        =   15
         Top             =   690
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
            Caption         =   "ǰ     ��"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1095
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   2
         Left            =   3840
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   690
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Left            =   4500
         TabIndex        =   19
         Top             =   390
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "��뱸��"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   2
         Left            =   4500
         TabIndex        =   26
         Top             =   60
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "������ȣ"
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   45
            Width           =   1065
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   0
         Left            =   8250
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   60
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   690
         Left            =   8790
         TabIndex        =   29
         Top             =   120
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   1217
         _Version        =   196609
         Caption         =   "        �˻�"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   4
         Left            =   90
         TabIndex        =   30
         Top             =   60
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "��������"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   1590
         TabIndex        =   31
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   116785153
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   2
         Left            =   2880
         TabIndex        =   32
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   116785153
         CurrentDate     =   36871
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   11850
      TabIndex        =   6
      Tag             =   "PERM_ADDNEW"
      Top             =   8670
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      �μ�(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13590
      TabIndex        =   7
      Top             =   8670
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      �ݱ�(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   480
      Index           =   1
      Left            =   7080
      TabIndex        =   24
      Top             =   8160
      Visible         =   0   'False
      Width           =   4080
      _cx             =   7197
      _cy             =   847
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
Attribute VB_Name = "frmSubulOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************
' �����̷�
'-----------------------------------------------------------------------------------------------------
'��ûID : S_201211_��������_03
'��û���� : 2012.11.22
'��û���� : ���Ҹ��� ������ ��µǰ�
'���泻�� : ���� ������� ����-���� �׸����μ�
'
'******************************************************************************************************
Option Explicit

Private m_bloading As Boolean
Dim sPrinter As String

'S_201211_��������_03 �� ���� �߰�
Private Const REPORTFILE = "\Report\����SubulReportOrder.xls"           '�Ϲݾ��

'���������� KG���� ����
''Private Const REPORTFILE_KG = "\Report\����SubulReportOrder_Kg.xls"     'KG���
'Private Const REPORTFILE = "\Report\SubulReport.rpt"                   '��ü ���� ��¿�  , S_201202_����_01�� ���� remark
'Private Const REPORTFILE2 = "\Report\SubulReportOneCust.rpt"           '1�� ��ü ��¿�   , S_201202_����_01�� ���� remark



'''S_201211_��������_03 �� ���� ����-�ּ�ó��
''Private Sub cmdPrint_Click()
''    pnlPrn.Visible = True
''End Sub


'S_201211_��������_03 �� ���� ����-NEW�ҽ�
Private Sub cmdPrint_Click()
    If grdData(0).Rows = grdData(0).FixedRows Then Exit Sub

    '------------------------------------------------------------------
    'S_201202_����_01 �� ���� Remark
    '------------------------------------------------------------------
    '    With grdData
    '    If optGub(0).Value = True Then
    '        If Trim(txtSearch(1).Tag) = "" And grdData.TextMatrix(grdData.Row, .ColIndex("CUSTOMID")) = "" Then        '�ŷ�ó �ڵ�
    '            MsgBox "�ŷ�ó�� ������ �� �μ��Ͻʽÿ�.", vbOKOnly
    '            Exit Sub
    '        End If
    '    End If
    '    End With
    '
    '    Call ReportPrint
    '------------------------------------------------------------------

    On Error GoTo ErrHandler

''    If Len(txtSearch(1).Tag) = 0 Then
''        MsgBox "���� ������ �ŷ�ó�� �������Ŀ� ������ �˴ϴ�." & vbCrLf & "���� �ŷ�ó�� �����Ͽ��ֽʽÿ�.", vbOKOnly
''        Exit Sub
''    End If

    Me.PopupMenu PlusMDI.mnuPopup

    Screen.MousePointer = vbHourglass

''    'KG ���� ����
''    If chkKG.Value = vbChecked Then     'KG���� ���
''        Call MakeExcelSubulReport(True)
''    Else
        Call MakeExcelSubulReport(False)
''    End If

    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmSubulOrder.cmdPrint_Click", Err.Description)

End Sub

Private Sub cmdPrnCancel_Click()
    pnlPrn.Visible = False
End Sub

Private Sub cmdPrnOK_Click()
    Dim II%, vCustom As Variant
    
    If optPrn(0).Value = True Then
        Call FillGrdList
    Else
        If cboCustom.Text = AllStr Then
           
            For II = 1 To cboCustom.ListCount - 1
                Call SetDataToPrn(cboCustom.List(II))
                
            Next II
        Else
            Call SetDataToPrn(cboCustom.Text)
        End If
    End If
    pnlPrn.Visible = False
    
End Sub

Sub FillGrdList()
    
    Dim sDate As String, eDate As String
    
    Dim i As Long, nRows As Long, II As Long, JJ As Long
       
    With grdData(1)
        .Rows = grdData(0).FixedRows
        .Cols = grdData(0).Cols
        .FixedRows = grdData(0).FixedRows
        .Redraw = flexRDBuffered
        .ExtendLastCol = False

        .GridLines = flexGridInset

        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        
        .FontSize = 9
        .FontName = "����"
        
        nRows = 0
        .Cell(flexcpText, nRows, 0, nRows, .Cols - 1) = "Order�� �� �� ��Ȳ"
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = True
        .RowHeight(nRows) = 800
        
        nRows = 1
        .RowHeight(nRows) = 500
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "�� �� �� ó : ���ŷ�ó "
        
        nRows = 2
        .RowHeight(nRows) = 500
        
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "�� �������� : " & MakeDate(DF_FULL, dtpDate(1)) & " ~ " & MakeDate(DF_FULL, dtpDate(2))
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        For i = 0 To .FixedRows - 1
           .MergeRow(i) = True
        Next i
        
        .Cell(flexcpAlignment, 1, 0, 2, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpBackColor, 3, 1, 4, .Cols - 1) = &HF5F5F5

        .ExtendLastCol = False
        .Redraw = flexRDDirect
        
        
        .ColHidden(0) = True
        .ColHidden(6) = True
        .ColHidden(9) = True
        
        .ColHidden(1) = False
        
        nRows = .Rows
        For II = grdData(0).FixedRows To grdData(0).Rows - 1
                .AddItem ""
                For JJ = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, JJ) = grdData(0).TextMatrix(II, JJ)
                Next JJ
                .RowHidden(.Rows - 1) = grdData(0).RowHidden(II)
        Next II
        
        .ScrollBars = flexScrollBarBoth
        .MergeCells = flexMergeFree
        .ExtendLastCol = False
        .RowHeightMin = 400
        
        Call SetPrintMode(grdData(1), 2, True)

        .PrintGrid "��������", True, 2, 100, 1000

        Call SetPrintMode(grdData(1), 2, False)
        
    End With

End Sub

Sub SetDataToPrn(ByVal kCustom As String)
    Dim II%, JJ%, sRows As Integer
    
    Call SetPrintMode(grdData(1), 5, True)
    Call FillGrdPrintHeader(kCustom)
    With grdData(1)
        sRows = .Rows
        For II = grdData(0).FixedRows To grdData(0).Rows - 1
            If grdData(0).TextMatrix(II, 1) = kCustom Then
                .AddItem ""
                For JJ = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, JJ) = grdData(0).TextMatrix(II, JJ)
                Next JJ
                .RowHidden(.Rows - 1) = grdData(0).RowHidden(II)
            End If
        Next II
        
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(6) = True
        .ColHidden(9) = True
        
        For II = 0 To .Cols - 1
            .Cell(flexcpAlignment, sRows, II, .Rows - 1, II) = grdData(0).ColAlignment(II)
        Next II
        
        .SheetBorder = vbBlack
        
        Call SetPrintMode(grdData(1), 2, True)
        
        .RowHeightMin = 400
        .PrintGrid "��������", True, 2, 100, 1000
        
        Call SetPrintMode(grdData(1), 2, False)
        
        .ColHidden(1) = False
    End With
End Sub






Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Load()
    Dim i%

    Me.Move 0, 0, 15360, 9840
    

    Call SetOperate(Me)
    Call ChangeMode(Me, True)
    
    Call InitGrid(0)
    Call InitGrid(1)
    
    dtpDate(1) = Now
    dtpDate(2) = Now
    
    pnlPrn.Visible = False
    
    With CboOrderFlag
        .AddItem "9.��ü"
        .AddItem "1.���"
        .AddItem "0.����"
        .ListIndex = 0
    End With
    
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(1).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    cmdFind(2).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(2).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    txtSearch(1).Enabled = False
    cmdFind(1).Enabled = False
    
    txtSearch(2).Enabled = False
    cmdFind(2).Enabled = False
    
End Sub

Sub FillGrdPrintHeader(ByVal kCustom As String)
    Dim i%, nRows As Integer
    Dim sDate As String, eDate As String
    
    With grdData(1)
        .Rows = grdData(0).FixedRows
        .FixedRows = grdData(0).FixedRows
        .Redraw = flexRDBuffered
        .ExtendLastCol = False

        .GridLinesFixed = flexGridNone
        .GridLines = flexGridInset
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        .RowHidden(4) = False
        
        
        nRows = 0
        .RowHeight(nRows) = 500
        .FontSize = 10
        
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "Order�� ���Ҹ���"
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 16
        .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = True
        .RowHeight(nRows) = 800
        
        
        nRows = 1
        .RowHeight(nRows) = 500
        
        .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = False
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "�� �� �� ó : " & Trim(kCustom)
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 11
        
        
        nRows = 2
        .RowHeight(nRows) = 500
        
        .Cell(flexcpFontBold, nRows, 0, nRows, .Cols - 1) = False
        .Cell(flexcpText, nRows, 1, nRows, .Cols - 1) = "�� �������� : " & MakeDate(DF_FULL, dtpDate(1)) & " ~ " & MakeDate(DF_FULL, dtpDate(2))
        .Cell(flexcpFontSize, nRows, 0, nRows, .Cols - 1) = 11
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        
        For i = 0 To 2
           .MergeRow(i) = True
        Next i
        
        .Cell(flexcpAlignment, 1, 0, 2, .Cols - 1) = flexAlignCenterCenter
        '.Cell(flexcpBackColor, 3, 2, 4, .Cols - 1) = &HE0E0E0
        .Cell(flexcpFontBold, 3, 2, 4, .Cols - 1) = True
        
        .ColWidth(2) = grdData(0).ColWidth(2) + 500
        .ColWidth(14) = 800
        .ColWidth(15) = 0
        
        .RowHeight(3) = 450
        .RowHeight(4) = 450
        .ExtendLastCol = True
        
        .MergeRow(.Rows - 2) = True
        .MergeCells = flexMergeFree
        
        '--- ���� ������ �κа� Merge �и��ϱ� ���� ����� �ϳ� ����
        .AddItem ""
        .RowHidden(.Rows - 1) = True
        .SheetBorder = vbBlack
        
        .GridLines = flexGridInset

        .ExtendLastCol = False
        .Redraw = flexRDDirect
    End With
    
End Sub


Private Sub chkSearch_Click(Index As Integer)
    Select Case Index
        Case 1
            txtSearch(1).Enabled = chkSearch(Index).Value
            cmdFind(1).Enabled = chkSearch(Index).Value
            If chkSearch(Index).Value Then
                txtSearch(1).SetFocus
            End If
        Case 2
            txtSearch(2).Enabled = chkSearch(2).Value
            cmdFind(2).Enabled = chkSearch(2).Value
            If chkSearch(2).Value Then
                txtSearch(2).SetFocus
            End If
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub







Private Sub txtSearch_GotFocus(Index As Integer)
    Call GotFocusText(txtSearch(Index))
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            Call ReturnCode(LG_CUSTOM, , False, txtSearch(Index))
        ElseIf Index = 2 Then
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
        End If
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 1
            Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
        Case 2
            Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
    End Select
End Sub

Private Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid(ByVal Index As Integer)
    Dim i%, nRows%

    Call SetVSFlexGrid(grdData(Index))
    With grdData(Index)
        .Rows = 5
        .Cols = 20
        
        .FixedRows = 5
        .FixedCols = 1
        
        .RowHeight(0) = 250
        .RowHeight(1) = 250

        nRows = 3
        .TextMatrix(nRows, 0) = " "
        .TextMatrix(nRows, 1) = "�ŷ�ó"
        .TextMatrix(nRows, 2) = "ǰ    ��"
        .TextMatrix(nRows, 3) = "����"
        .TextMatrix(nRows, 4) = "��������"
        .TextMatrix(nRows, 5) = "�����̿�"
        .TextMatrix(nRows, 6) = "�԰�"
        .TextMatrix(nRows, 7) = "�԰�"
        .TextMatrix(nRows, 8) = ""
        .TextMatrix(nRows, 9) = "���"
        .TextMatrix(nRows, 10) = "���"
        .TextMatrix(nRows, 11) = "���"
        .TextMatrix(nRows, 12) = ""
        .TextMatrix(nRows, 13) = "������"
        .TextMatrix(nRows, 14) = "���"
        .TextMatrix(nRows, 15) = "������ȣ"
        .TextMatrix(nRows, 16) = "OrderNO"
        .TextMatrix(nRows, 17) = "Memo"
        .TextMatrix(nRows, 18) = "Cls"
        .TextMatrix(nRows, 19) = "pkey"
        
        
        nRows = 4
        '�ŷ�ó, ǰ��, OrderID, OrderNO, �԰�����, �����ó, ����, ����, �������,����, �ҿ䷮, ���, ���, �޸�
        .TextMatrix(nRows, 0) = " ":                 .ColWidth(0) = 0
        .TextMatrix(nRows, 1) = "�ŷ�ó":            .ColWidth(1) = 1500:     .ColAlignment(1) = flexAlignLeftCenter:    .FixedAlignment(1) = flexAlignCenterCenter
        .TextMatrix(nRows, 2) = "ǰ    ��":          .ColWidth(2) = 2000:     .ColAlignment(2) = flexAlignLeftCenter:    .FixedAlignment(2) = flexAlignCenterCenter
        .TextMatrix(nRows, 3) = "����":              .ColWidth(3) = 700:      .ColAlignment(3) = flexAlignCenterCenter:  .FixedAlignment(3) = flexAlignCenterCenter
        .TextMatrix(nRows, 4) = "��������":          .ColWidth(4) = 1100:     .ColAlignment(4) = flexAlignLeftCenter:    .FixedAlignment(4) = flexAlignCenterCenter
        .TextMatrix(nRows, 5) = "�����̿�":          .ColWidth(5) = 1200:     .ColAlignment(5) = flexAlignRightCenter:   .FixedAlignment(5) = flexAlignCenterCenter
        .TextMatrix(nRows, 6) = "����":              .ColWidth(6) = 900:      .ColAlignment(6) = flexAlignRightCenter:   .FixedAlignment(6) = flexAlignCenterCenter
        .TextMatrix(nRows, 7) = "����":              .ColWidth(7) = 1200:     .ColAlignment(7) = flexAlignRightCenter:   .FixedAlignment(7) = flexAlignCenterCenter
        .TextMatrix(nRows, 8) = "":                  .ColWidth(8) = 0:        .ColAlignment(8) = flexAlignCenterCenter:  .FixedAlignment(8) = flexAlignCenterCenter
        .TextMatrix(nRows, 9) = "����":              .ColWidth(9) = 900:      .ColAlignment(9) = flexAlignRightCenter:   .FixedAlignment(9) = flexAlignCenterCenter
        .TextMatrix(nRows, 10) = "���":           .ColWidth(10) = 1200:    .ColAlignment(10) = flexAlignRightCenter:  .FixedAlignment(10) = flexAlignCenterCenter
        .TextMatrix(nRows, 11) = "�ҿ䷮":           .ColWidth(11) = 1300:    .ColAlignment(11) = flexAlignRightCenter:  .FixedAlignment(11) = flexAlignCenterCenter
        .TextMatrix(nRows, 12) = "":                 .ColWidth(12) = 0:       .ColAlignment(12) = flexAlignRightCenter:  .FixedAlignment(12) = flexAlignCenterCenter
        .TextMatrix(nRows, 13) = "������":           .ColWidth(13) = 1300:    .ColAlignment(13) = flexAlignRightCenter:  .FixedAlignment(13) = flexAlignCenterCenter
        .TextMatrix(nRows, 14) = "���":             .ColWidth(14) = 0:       .ColAlignment(14) = flexAlignLeftCenter:   .FixedAlignment(14) = flexAlignCenterCenter
        .TextMatrix(nRows, 15) = "������ȣ":         .ColWidth(15) = 0:       .ColAlignment(15) = flexAlignCenterCenter: .FixedAlignment(15) = flexAlignCenterCenter
        .TextMatrix(nRows, 16) = "OrderNO":          .ColWidth(16) = 2400:    .ColAlignment(16) = flexAlignLeftCenter:   .FixedAlignment(16) = flexAlignCenterCenter
        .TextMatrix(nRows, 17) = "Memo":             .ColWidth(17) = 0:       .ColAlignment(17) = flexAlignCenterCenter: .FixedAlignment(17) = flexAlignCenterCenter
        .TextMatrix(nRows, 18) = "Cls":              .ColWidth(18) = 0
        .TextMatrix(nRows, 19) = "pkey":             .ColWidth(19) = 0
        
        .ColKey(1) = "CUSTOM":          .ColKey(2) = "ARTICLE":     .ColKey(3) = "DATE"
        .ColKey(4) = "WORKNAME":        .ColKey(5) = "PREMONTHQTY": .ColKey(6) = "INROLLQTY"
        .ColKey(7) = "INQTY":           .ColKey(8) = "":            .ColKey(9) = "OUTROLLQTY"
        .ColKey(10) = "OUTQTY":         .ColKey(11) = "OUTREALQTY": .ColKey(12) = ""
        .ColKey(13) = "OVERQTY":        .ColKey(14) = "REMARK":     .ColKey(15) = "ORDERID"
        .ColKey(16) = "ORDERNO":        .ColKey(17) = "MEMO":       .ColKey(18) = "CLS"
        .ColKey(19) = "PKEY"
        
        .Cell(flexcpFontBold, 3, 0, 4, .Cols - 1) = True
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(3) = True
        
        
        For i = 0 To .FixedRows - 3
            .RowHidden(i) = True
        Next i
        
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next i
        
        .ExtendLastCol = True
        .ExplorerBar = flexExNone
        .Editable = flexEDKbdMouse
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect
        
    End With

End Sub

Private Sub FillGridData()
    Dim oSubul As PlusLib2.CSubul
    Dim rs       As Recordset, dCustom_str As String, dArticle_Str As String, dDate_str As String
    Dim sOrderID$, bFlag As Boolean, II%, dCustom$
    Dim i As Long

    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler

    m_bloading = True

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon

    Set rs = oSubul.GetSubulOrder("1", MakeDate(DF_SHORT, dtpDate(1)), MakeDate(DF_SHORT, dtpDate(2)) _
                        , IIf(chkSearch(0).Value = vbChecked, 1, 0), txtSearch(0).Text _
                        , IIf(chkSearch(1).Value = vbChecked, 1, 0), txtSearch(1).Tag _
                        , IIf(chkSearch(2).Value = vbChecked, 1, 0), txtSearch(2).Tag _
                        , Left(CboOrderFlag, 1))
    Set oSubul = Nothing
    cboCustom.Clear
    cboCustom.AddItem AllStr
    With grdData(0)
        .Redraw = flexRDNone

        .Rows = .FixedRows
        For i = 1 To rs.RecordCount
        
            If Trim(dCustom) = "" Then
                dCustom = rs!kCustom
                cboCustom.AddItem dCustom
            ElseIf Trim(dCustom) <> Trim(rs!kCustom) Then
                    cboCustom.AddItem dCustom
            End If
            
            .AddItem CStr(i) & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article) & vbTab & _
                     MakeDate(DF_MD, rs!IODate) & vbTab & rs!WorkName & vbTab & SetCurrency(rs!BeforeQty) & vbTab & _
                     IIf(rs!StuffRoll = 0, "", rs!StuffRoll) & vbTab & IIf(rs!StuffQty = 0, "", SetCurrency(rs!StuffQty, 0)) & vbTab & "" & vbTab & _
                     IIf(rs!OutRoll = 0, "", rs!OutRoll) & vbTab & IIf(rs!OutQty = 0, " ", SetCurrency(rs!OutQty, 0)) & "" & vbTab & _
                     IIf(rs!OutRealQty = 0, "", SetCurrency(rs!OutRealQty, 0)) & vbTab & "" & vbTab & SetCurrency(rs!AfterQty) & vbTab & "" & vbTab & _
                     MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & Trim(rs!OrderNo) & vbTab & "" & vbTab & rs!Cls
            
            Select Case rs!Cls
                Case "0"
                    .TextMatrix(.Rows - 1, 3) = ""
                    .TextMatrix(.Rows - 1, 4) = "�����̿�"
                Case "3"
'                    .TextMatrix(.Rows - 1, 13) = SetCurrency(rs!StuffQty - rs!OutRealQty, 0)
                    .TextMatrix(.Rows - 1, 3) = ""
                    .TextMatrix(.Rows - 1, 4) = "�Ұ�"
                    .TextMatrix(.Rows - 1, 16) = ""
                    .Cell(flexcpFontBold, .Rows - 1, 3, .Rows - 1, .Cols - 1) = True
            End Select
            
            If rs!nCount = 1 Then
                .RowHidden(.Rows - 1) = True
            End If

                     
            .AddItem CStr(i) & vbTab & Trim(rs!kCustom) & vbTab & Trim(rs!Article)
            .RowHidden(.Rows - 1) = True
            
            dCustom_str = Trim(rs!kCustom)
            dArticle_Str = Trim(rs!Article)
            dDate_str = Trim(rs!IODate)
            
            rs.MoveNext
        Next i
        .Redraw = flexRDDirect
        
        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If
        
        .MergeCells = flexMergeFree
        
        .Redraw = flexRDDirect
        .SetFocus
        cboCustom.ListIndex = 0
    End With
    
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault

    Set rs = Nothing
    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmSubulOrder.FillGridData", Err.Description)
End Sub

'S_201211_��������_03 �� ���� �߰�
Private Sub MakeExcelSubulReport(pbKG As Boolean)
    Dim oExcel                          As Excel.Application
    Dim oExcelBook                      As Excel.Workbook
    Dim oExcelSheet                     As Excel.Worksheet
    Dim oRange                          As Excel.Range
    Dim oFs                             As FileSystemObject
    Dim oCustom                         As PlusLib2.CCustom
    Dim oOutware                        As PlusLib2.COutWare
    Dim rs                              As ADODB.Recordset
    Dim lssql                           As String
    Dim lstempQty                       As String
    Dim lsTempLossQty                   As String
    Dim i%, j%, nRow%, nLimitRow%, nBaseRow%, nCurRow%, nPage%, sReport$, sReport1$
    Dim nOrderSeq%, sLotNo$
    Dim nColorRoll%, nColorQty#, nColorLossQty#
    Dim sWorkWidth                      As String
    Dim EXCEL_1PageData_ROW             As Integer
    Dim vColorSum()                     As Double
    Dim sUnit                           As String
    Dim nSeq                            As Integer
    Dim sDate                           As String
    Dim sArticle                      As String
    Dim sOrderNO                        As String
    Dim sCustomID                       As String       '�ŷ�ó �ڵ�
            
    Dim bAllPrint         As Boolean              '�ŷ�ó��ü ������� üũ
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
''    '���������� KG���� ����
''    If pbKG = True Then
'''        sReport1 = App.Path & "\Report\tmpSubulReportOrder_kg.xls"
''        sReport = App.Path & REPORTFILE_KG
''    Else
'        sReport1 = App.Path & "\Report\tmpSubulReportOrder.xls"
        sReport = App.Path & REPORTFILE
''    End If
    
    Set oExcel = New Excel.Application
    '���� ����open
    Set oExcelBook = oExcel.Workbooks.Open(sReport)
    
''    '//������ �Ʒ� �ּ� ����
''    oExcel.WindowState = xlMaximized
''    oExcel.Application.Visible = True
 
    EXCEL_1PageData_ROW = 47
    '---------------------------------------------
    
    
    If chkSearch(1).Value = 0 Or Len(txtSearch(1).Tag) = 0 Then
        bAllPrint = True        '�ŷ�ó ��ü ���
    End If
    With oExcel
       ' Make Sum
       
       If bAllPrint = True Then         ' �ŷ�ó ��ü�ΰ��
            .Worksheets("Form2").Activate
       Else
            .Worksheets("Form").Activate
       End If
        
      
        .Cells(2, 1) = "��" & g_companyInfo.Company_Name
        .Cells(3, 6) = "[����Ⱓ " & MakeDate(DF_FULL, dtpDate(1)) & " - " & MakeDate(DF_FULL, dtpDate(2)) & "] "      '[����Ⱓ ]
        
        If bAllPrint = True Then         ' �ŷ�ó ��ü�ΰ��
            .Cells(4, 6) = " (��ü)"                                                        '�ŷ�ó
        Else
            .Cells(4, 6) = txtSearch(1).Tag & " -�� " & txtSearch(1)                                                        '�ŷ�ó
        End If
        
        If chkSearch(2).Value = 0 Or Len(txtSearch(2).Tag) = 0 Then            ' �ŷ�ó ��ü�ΰ��
            .Cells(5, 6) = " (��ü)"                                                        '�ŷ�ó
        Else
            .Cells(5, 6) = IIf(chkSearch(2).Value = True And txtSearch(2).Text <> "", txtSearch(2).Text, "")                'ǰ��
        End If
        
        .Worksheets("Print").Activate
        
        nPage = 1
        nBaseRow = GetExcelRollBaseRow(nPage, EXCEL_1PageData_ROW)
        
        Call InsertExcelForm(oExcel, nPage, EXCEL_1PageData_ROW, IIf(bAllPrint = True, 0, 1))       '��ü ����� ��� 0, �ŷ�ó1���� ��� 1
        
        nCurRow = nBaseRow + 8
            
        nOrderSeq = 0
        For i = grdData(0).FixedRows To grdData(0).Rows - 1  ' Step 2
        
            If grdData(0).RowHidden(i) = True Or grdData(0).TextMatrix(i, grdData(0).ColIndex("CLS")) = "9" Then GoTo Next_i '"9": ��ü ��
            If nCurRow + nRow > nBaseRow + EXCEL_1PageData_ROW Then
                nPage = nPage + 1
                nBaseRow = GetExcelRollBaseRow(nPage, EXCEL_1PageData_ROW)
                 Call InsertExcelForm(oExcel, nPage, EXCEL_1PageData_ROW, IIf(bAllPrint = True, 0, 1))       '��ü ����� ��� 0, �ŷ�ó1���� ��� 1
                nCurRow = nBaseRow + 8
                nRow = 0
                sOrderNO = ""                                                       'OrderNo
                sDate = ""                                                          '����
                sArticle = ""                                                     'ǰ��
                sCustomID = ""
            
                If bAllPrint = True Then           ' �ŷ�ó ��ü�ΰ��
                    .Cells(nCurRow + nRow, 1) = grdData(0).TextMatrix(i, grdData(0).ColIndex("CUSTOM"))            '�ŷ�ó
                End If
                
            End If
            
            If sCustomID = "" Then
                sCustomID = grdData(0).TextMatrix(i, grdData(0).ColIndex("CUSTOM"))
                .Cells(nCurRow + nRow, 1) = grdData(0).TextMatrix(i, grdData(0).ColIndex("CUSTOM"))            '�ŷ�ó
            ElseIf sCustomID <> grdData(0).TextMatrix(i, grdData(0).ColIndex("CUSTOM")) Then
                sCustomID = grdData(0).TextMatrix(i, grdData(0).ColIndex("CUSTOM"))
                .Cells(nCurRow + nRow, 1) = grdData(0).TextMatrix(i, grdData(0).ColIndex("CUSTOM"))            '�ŷ�ó
            Else
                .Cells(nCurRow + nRow, 1) = ""                                                                  '�ŷ�ó
            End If
            
            If bAllPrint = True Then           ' �ŷ�ó ��ü�ΰ��
                .Cells(nCurRow + nRow, 6) = grdData(0).TextMatrix(i, grdData(0).ColIndex("ARTICLE"))           'ǰ��
            Else
                .Cells(nCurRow + nRow, 1) = grdData(0).TextMatrix(i, grdData(0).ColIndex("ARTICLE"))           'ǰ��
            End If
            
            .Cells(nCurRow + nRow, 12) = grdData(0).TextMatrix(i, grdData(0).ColIndex("DATE"))              '����
            .Cells(nCurRow + nRow, 14) = grdData(0).TextMatrix(i, grdData(0).ColIndex("WORKNAME"))        '��������
            .Cells(nCurRow + nRow, 18) = grdData(0).TextMatrix(i, grdData(0).ColIndex("PREMONTHQTY"))     '�����̿�

            .Cells(nCurRow + nRow, 22) = grdData(0).TextMatrix(i, grdData(0).ColIndex("INROLLQTY"))       '�԰�����
            .Cells(nCurRow + nRow, 25) = grdData(0).TextMatrix(i, grdData(0).ColIndex("INQTY"))           '�԰����
            
            .Cells(nCurRow + nRow, 29) = grdData(0).TextMatrix(i, grdData(0).ColIndex("OUTROLLQTY"))      '�������
            .Cells(nCurRow + nRow, 32) = grdData(0).TextMatrix(i, grdData(0).ColIndex("OUTQTY"))          '������
            .Cells(nCurRow + nRow, 36) = grdData(0).TextMatrix(i, grdData(0).ColIndex("OUTREALQTY"))      '���ҿ䷮
            
            .Cells(nCurRow + nRow, 40) = grdData(0).TextMatrix(i, grdData(0).ColIndex("OVERQTY"))         '������
            .Cells(nCurRow + nRow, 44) = grdData(0).TextMatrix(i, grdData(0).ColIndex("ORDERNO"))         'OrderNo
                            
            sDate = grdData(0).TextMatrix(i, grdData(0).ColIndex("DATE"))                                 '����
            sArticle = grdData(0).TextMatrix(i, grdData(0).ColIndex("ARTICLE"))                       'ǰ��ID
            nRow = nRow + 1
Next_i:
        
        
        Next i
        
        If nCurRow + nRow > nBaseRow + EXCEL_1PageData_ROW Then
            nPage = nPage + 1
            nBaseRow = GetExcelRollBaseRow(nPage, EXCEL_1PageData_ROW)
             Call InsertExcelForm(oExcel, nPage, EXCEL_1PageData_ROW, IIf(bAllPrint = True, 0, 1))       '��ü ����� ��� 0, �ŷ�ó1���� ��� 1
            nCurRow = nBaseRow + 7
            nRow = 0
            
            If bAllPrint = True Then           ' �ŷ�ó ��ü�ΰ��
                .Cells(nCurRow + nRow, 1) = grdData(0).TextMatrix(i, grdData(0).ColIndex("CUSTOM"))            '�ŷ�ó
            End If

        End If
    
    End With


   '---------------------------------------------
   '������ Report FIle �� ������ �����ϰ�
   '---------------------------------------------
    Set oFs = New FileSystemObject
    '���Ҹ��� ���� ���� ��� ����
    If Not oFs.FolderExists(CStr(App.Path) & "\���Ҹ���\") Then
        oFs.CreateFolder (CStr(App.Path) & "\���Ҹ���\")           '������� ���� ����
    End If

    'KG ���Ҿ���
''    If pbKG = True Then     'kg ���Ҹ���
''''        sReport1 = App.Path & "\Report\tmpSubulReportOrder_kg.xls"
''        sReport1 = App.Path & "\���Ҹ���\���������Ҹ���_kg_" & Left(MakeDate(DF_SHORT, dtpDate(1)), 6) & "_" & txtSearch(1) & ".xls"
''
''    Else
''        sReport1 = App.Path & "\Report\tmpSubulReportOrder.xls"
        sReport1 = App.Path & "\���Ҹ���\���������Ҹ���_" & Left(MakeDate(DF_SHORT, dtpDate(1)), 6) & "_" & txtSearch(1) & ".xls"
    
''    End If
    
    If oFs.FileExists(sReport1) Then Call oFs.DeleteFile(sReport1)
''    oFs.CopyFile sReport, sReport1
    Set oFs = Nothing
    
''    Call oExcelBook.Save
    Call oExcelBook.SaveAs(sReport1)
    
    oExcel.WindowState = xlMaximized
    oExcel.Application.Visible = True
    oExcel.ActiveWindow.SelectedSheets.PrintPreview

    Screen.MousePointer = vbDefault

    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    Set oFs = Nothing
    
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[MakeExcelSubulReport]"
    End If
       
    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    Set oFs = Nothing

End Sub

 
'S_201211_��������_03 �� ���� �߰�
Private Function GetExcelRollBaseRow(nPage, li1PageRow As Integer)
    GetExcelRollBaseRow = (nPage - 1) * li1PageRow
End Function

'S_201211_��������_03 �� ���� �߰�
Private Function InsertExcelForm(oExcel As Excel.Application, nPage As Integer, li1PageRow As Integer, nPrnGub As Integer)
    Dim i%, nBaseRow%

    nBaseRow = GetExcelRollBaseRow(nPage, li1PageRow)
    With oExcel
    
        If nPrnGub = 1 Then
            .Sheets("Form").Select          '�ŷ�ó1�� ���
        Else
            .Sheets("Form2").Select         '�ŷ�ó ��ü ���
        End If
        

        .Rows("1:" & CStr(li1PageRow)).Select
        .Selection.Copy

        .Sheets("Print").Select
        .Rows(CStr(nBaseRow + 1) & ":" & CStr(nBaseRow + 1)).Select
        .Selection.Insert Shift:=xlDown
        
      '  .Cells(nBaseRow + 6, 5) = "PAGE : " & nPage
    End With
End Function

