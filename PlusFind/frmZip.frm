VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmZip 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�ּ� ã��"
   ClientHeight    =   8175
   ClientLeft      =   2910
   ClientTop       =   1605
   ClientWidth     =   7245
   ClipControls    =   0   'False
   Icon            =   "frmZip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin TabDlg.SSTab SSTab1 
      Height          =   7995
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   14102
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "���θ�"
      TabPicture(0)   =   "frmZip.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstZip(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdClose(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSelect(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmZip"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "����"
      TabPicture(1)   =   "frmZip.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstZip(1)"
      Tab(1).Control(1)=   "cmdClose(1)"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "cmdSelect(1)"
      Tab(1).ControlCount=   4
      Begin VB.Frame frmZip 
         Height          =   7155
         Left            =   180
         TabIndex        =   13
         Top             =   345
         Width           =   6705
         Begin VB.Frame fraDetail 
            BorderStyle     =   0  '����
            Enabled         =   0   'False
            Height          =   795
            Index           =   0
            Left            =   60
            TabIndex        =   31
            Top             =   6300
            Width           =   6615
            Begin VB.TextBox txtGunMoolMngNo 
               Height          =   345
               Index           =   0
               Left            =   0
               TabIndex        =   36
               Top             =   390
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtAssist 
               Height          =   345
               Index           =   0
               Left            =   1110
               TabIndex        =   35
               Top             =   390
               Width           =   5415
            End
            Begin VB.TextBox txtZipCode 
               Height          =   345
               Index           =   0
               Left            =   0
               TabIndex        =   34
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txtAddr 
               Height          =   345
               Index           =   0
               Left            =   1110
               TabIndex        =   33
               Top             =   0
               Width           =   4005
            End
            Begin VB.TextBox txtAddr2 
               Height          =   345
               Index           =   0
               Left            =   5130
               TabIndex        =   32
               Top             =   0
               Width           =   1395
            End
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "�˻�"
            Height          =   345
            Index           =   0
            Left            =   4740
            TabIndex        =   24
            Top             =   2310
            Width           =   915
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Index           =   0
            Left            =   1470
            TabIndex        =   23
            Top             =   2310
            Width           =   3135
         End
         Begin VB.Frame fraSearch 
            Height          =   1215
            Left            =   90
            TabIndex        =   14
            Top             =   1050
            Width           =   6435
            Begin VB.OptionButton optSearch 
               Caption         =   "���θ�+�ǹ���ȣ"
               Height          =   285
               Index           =   0
               Left            =   180
               TabIndex        =   20
               Top             =   270
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.OptionButton optSearch 
               Caption         =   "��(��/��/��)��+����"
               Height          =   285
               Index           =   1
               Left            =   2100
               TabIndex        =   19
               Top             =   270
               Width           =   2025
            End
            Begin VB.OptionButton optSearch 
               Caption         =   "�ǹ���(����Ʈ ��)"
               Height          =   285
               Index           =   2
               Left            =   4380
               TabIndex        =   18
               Top             =   270
               Width           =   1785
            End
            Begin VB.OptionButton optSearch 
               Caption         =   "�缭��+�缭�Թ�ȣ"
               Height          =   315
               Index           =   3
               Left            =   5820
               TabIndex        =   17
               Top             =   450
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.ComboBox cboSiDo 
               Height          =   300
               Left            =   990
               Style           =   2  '��Ӵٿ� ���
               TabIndex        =   16
               Top             =   690
               Width           =   1485
            End
            Begin VB.ComboBox cboSiGunGu 
               Height          =   300
               Left            =   3270
               Style           =   2  '��Ӵٿ� ���
               TabIndex        =   15
               Top             =   690
               Width           =   2535
            End
            Begin VB.Label Label8 
               Alignment       =   1  '������ ����
               Caption         =   "�õ�"
               Height          =   270
               Left            =   240
               TabIndex        =   22
               Top             =   750
               Width           =   600
            End
            Begin VB.Label Label9 
               Alignment       =   1  '������ ����
               Caption         =   "�ñ��� "
               Height          =   270
               Left            =   2610
               TabIndex        =   21
               Top             =   750
               Width           =   600
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid grdZipList 
            Height          =   3645
            Index           =   0
            Left            =   60
            TabIndex        =   25
            Top             =   2640
            Width           =   6525
            _cx             =   11509
            _cy             =   6429
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
            Rows            =   2
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmZip.frx":0044
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
         Begin VB.Label Label3 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            BorderStyle     =   1  '���� ����
            Caption         =   "�� �� ��"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   30
            Top             =   2340
            Width           =   1320
         End
         Begin VB.Label Label2 
            Caption         =   "�� �˻����"
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   180
            Width           =   6360
         End
         Begin VB.Label Label5 
            Caption         =   "  - ���θ�(~��,~��)+�ǹ���ȣ            ��)���뱸��"
            Height          =   180
            Left            =   120
            TabIndex        =   28
            Top             =   420
            Width           =   6360
         End
         Begin VB.Label Label6 
            Caption         =   "  - ��(��/��/��)��                            ��)�湫��1�� 21-1"
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   630
            Width           =   6360
         End
         Begin VB.Label Label7 
            Caption         =   "  - �ǹ���(����Ʈ ��)                        ��)�����߾ӿ�ü��"
            Height          =   210
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   6360
         End
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "����"
         Height          =   360
         Index           =   0
         Left            =   5100
         TabIndex        =   12
         Top             =   7530
         Width           =   870
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "����"
         Height          =   360
         Index           =   1
         Left            =   -69900
         TabIndex        =   11
         Top             =   7530
         Width           =   870
      End
      Begin VB.Frame Frame2 
         Height          =   7125
         Left            =   -74790
         TabIndex        =   5
         Top             =   375
         Width           =   6675
         Begin VB.Frame fraDetail 
            BorderStyle     =   0  '����
            Enabled         =   0   'False
            Height          =   825
            Index           =   1
            Left            =   30
            TabIndex        =   37
            Top             =   6270
            Width           =   6585
            Begin VB.TextBox txtGunMoolMngNo 
               Height          =   345
               Index           =   1
               Left            =   30
               TabIndex        =   42
               Top             =   390
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox txtAssist 
               Height          =   345
               Index           =   1
               Left            =   960
               TabIndex        =   41
               Top             =   390
               Visible         =   0   'False
               Width           =   5535
            End
            Begin VB.TextBox txtZipCode 
               Height          =   345
               Index           =   1
               Left            =   30
               TabIndex        =   40
               Top             =   0
               Width           =   945
            End
            Begin VB.TextBox txtAddr 
               Height          =   345
               Index           =   1
               Left            =   960
               TabIndex        =   39
               Top             =   0
               Width           =   4125
            End
            Begin VB.TextBox txtAddr2 
               Height          =   345
               Index           =   1
               Left            =   5070
               TabIndex        =   38
               Top             =   0
               Width           =   1425
            End
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Index           =   1
            Left            =   1470
            TabIndex        =   7
            Top             =   1020
            Width           =   3135
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "�˻�"
            Height          =   345
            Index           =   1
            Left            =   4650
            TabIndex        =   6
            Top             =   990
            Width           =   915
         End
         Begin VSFlex7LCtl.VSFlexGrid grdZipList 
            Height          =   4845
            Index           =   1
            Left            =   60
            TabIndex        =   8
            Top             =   1380
            Width           =   6435
            _cx             =   11351
            _cy             =   8546
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
         Begin VB.Label Label4 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            BorderStyle     =   1  '���� ����
            Caption         =   "�� (��/��/��)"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   10
            Top             =   1050
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   $"frmZip.frx":0122
            Height          =   600
            Left            =   120
            TabIndex        =   9
            Top             =   300
            Width           =   5430
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ݱ�"
         Height          =   360
         Index           =   0
         Left            =   6000
         TabIndex        =   4
         Top             =   7530
         Width           =   870
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ݱ�"
         Height          =   360
         Index           =   1
         Left            =   -69000
         TabIndex        =   3
         Top             =   7530
         Width           =   870
      End
      Begin VB.ListBox lstZip 
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   7560
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.ListBox lstZip 
         Height          =   240
         Index           =   1
         Left            =   -74790
         TabIndex        =   1
         Top             =   7590
         Visible         =   0   'False
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'** System �� : MRRPLUS2-PlusFind
'** Author    : Wizard
'** �ۼ���    :
'** ����      : �ּ� �˻� ȭ��
'** ��������  :
'** ��������  : 2013.12.12
'**------------------------------------------------------------------------------------------------
'
'  ��û���� ID: S_201312_���_99
'  ��û��:
'  ���泯¥ : 2013.12.12
'  �۾���   : ���¿�
'  ��û���� : �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�
'  ���泻�� : ���θ� �ּ� �Է��� ���� �߰� �� ListBox  ��� �׸��� ���, ��� ��Ʈ�� �迭�� �ۼ���
'             index 0�� ���θ� �ּ� ���� ��Ʈ��, index 1 �� �� �����ּ� �˻��� ��Ʈ��
'**************************************************************************************************
Public ReturnStatus As Boolean

'S_201312_���_99 �� ���� �߰�
Public g_OldNNewClss As String          '0:���θ� �ּ�, 1:�� �����ּ�

'S_201312_���_99 �� ���� �߰�
Private Sub cboSiDo_Click()
    Call LS_MakeComboSiGunGu
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub cmdClose_Click(Index As Integer)
    Me.Hide
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub cmdFind_Click(Index As Integer)
    Call FilllstZip(Index)
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub FilllstZip(Index As Integer)
    Dim Query As String, iCount As Long
    Dim rs As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Dim sSearch(1) As String
    Dim lsAdditemStr                    As String
    Dim sSearchGub As String
    Dim sSiDo As String
    Dim sSiDTable As String
    Dim sSiGunGu As String
    
    If Len(txtName(Index)) = 0 Then
        MsgBox LoadResString(111), vbInformation
        Exit Sub
    End If

    'S_201312_���_99 �� ���� �߰�---------------------------------------
    sSiDo = Trim(Mid(cboSiDo.List(cboSiDo.ListIndex), 50))
''    sSiDTable = "ZipZoneAddress"
    Select Case sSiDo
        
        '����
        Case 11
            sSiDTable = "ZipZone_Seoul"
            
        '�λ�
        Case 26
            sSiDTable = "ZipZone_BuSan"
            
        '�뱸
        Case 27
            sSiDTable = "ZipZone_DaeGu"
            
        '��õ
        Case 28
            sSiDTable = "ZipZone_InCheon"
            
        '����
        Case 29
            sSiDTable = "ZipZone_GwangJu"
                    
        '����
        Case 30
            sSiDTable = "ZipZone_DaeJeon"
                    
        '���
        Case 31
            sSiDTable = "ZipZone_UlSan"
                    
        '����Ư����ġ��
        Case 36
            sSiDTable = "ZipZone_SeJong"
                    
        '��⵵
        Case 41
            sSiDTable = "ZipZone_GyeongGi"
        
        '������
        Case 42
            sSiDTable = "ZipZone_GangWon"
                    
        '��û�ϵ�
        Case 43
            sSiDTable = "ZipZone_ChungBuk"
                    
        '��û����
        Case 44
            sSiDTable = "ZipZone_ChungNam"
                    
        '����ϵ�
        Case 45
            sSiDTable = "ZipZone_JeonBuk"
                    
        '���󳲵�
        Case 46
            sSiDTable = "ZipZone_JeonNam"
                    
        '���ϵ�
        Case 47
            sSiDTable = "ZipZone_GyeongBuk"
            
        '��󳲵�
        Case 48
            sSiDTable = "ZipZone_GyeongNam"
                    
        '����Ư����ġ��
        Case 50
            sSiDTable = "ZipZone_JeJu"
    
    
    End Select
    
    sSiGunGu = Trim(Mid(cboSiGunGu.List(cboSiGunGu.ListIndex), 50))
    If optSearch(1).Value = True Then           '��(��/��/��)��+����
        sSearchGub = "1"
    ElseIf optSearch(2).Value = True Then       '�ǹ���(����Ʈ ��)
        sSearchGub = "2"
    Else                        '���θ�+�ǹ���ȣ
        sSearchGub = "0"
    End If

    sSearch(Index) = Trim(txtName(Index))
    If Index = 0 Then
        sSearch(Index) = Replace(Trim(txtName(Index)), " ", "")
    End If
    '---------------------------------------------------------------------
    
    Set adoCmd = New ADODB.Command

    If Index = 0 Then           '���θ� �ּ�
    
        'S_201312_���_99 �� ���� �߰�----------------------------------------------------------------
        With adoCmd
'            'S_201312_���_99 �� ���� ����(OLD:adoCon)
            .ActiveConnection = adoWizCon
            .CommandType = adCmdStoredProc
            'S_201312_���_99 �� ���� ����-OLD
''            .CommandText = "ZipDB.dbo.xp_ZipCode_sAddress"
            .CommandText = "xp_ZipCode_sAddress"
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Index)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sSearchGub)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, sSiDo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sSiDTable)
    ''        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkSiGunGu)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, sSiGunGu)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 51, sSearch(Index))
            Set rs = .Execute
        End With

        With grdZipList(0)
            .Rows = .FixedRows          '�׸��� �ʱ�ȭ

            Do While Not rs.EOF
                lsAdditemStr = CStr(.Rows)                                                                      '0)Row ��
                lsAdditemStr = lsAdditemStr & vbTab & Mid(rs!Zip_Code, 1, 3) & "-" & Mid(rs!Zip_Code, 4, 3)     '1)�����ȣ
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Addr)                                        '2)�⺻ �ּ�
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Addr2)                                       '3)�ǹ�����+�ι�
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!AssistAddr)                                  '4)���� �ּ�
                lsAdditemStr = lsAdditemStr & vbTab & ""                                                        '5)
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!GunMoolMng_No)                               '6)�ǹ������ĺ���ȣ
                
                .AddItem lsAdditemStr
    
                rs.MoveNext
            Loop
            rs.Close
    
            '�˻��ڷ� ������ ù�� ����
            If .Rows > .FixedRows Then
                .Row = 1
                .SetFocus
            Else
''                MsgBox "�˻��������", vbInformation, "�˻�"
                txtName(0).SetFocus
            End If
    
        End With
        '----------------------------------------------------------------------------------------------------
        
    
    '�� ���� �ּ�
    Else
        With adoCmd
            ''S_201312_���_99 �� ���� ����(OLD:adoCon)
            .ActiveConnection = adoWizCon
            .CommandType = adCmdStoredProc
            .CommandText = "xp_Common_sZipCode"
            .Parameters.Append .CreateParameter("xp_common_sZipCode", adVarChar, adParamInput, 51, sSearch(Index))
            Set rs = .Execute
        End With
        
''        With lstZip(Index)
''            .Clear
''
''            Do While Not rs.EOF
''                .AddItem rs!ZipCode & Space(3) & rs!City & " " & rs!Section & " " & _
''                        rs!Village & " " & Format(rs!Detail1) & " " & Format(rs!Detail2)
''
''                rs.MoveNext
''            Loop
''            rs.Close
''
''            If .ListCount > 0 Then
''                .ListIndex = 0
''                .SetFocus
''            Else
''                MsgBox LoadResString(111), vbInformation
''                txtName(Index).SetFocus
''            End If
''        End With

        'S_201312_���_99 �� ����  ����(ListBox���� Grid�� ����)
        With grdZipList(1)
            .Rows = .FixedRows          '�׸��� �ʱ�ȭ

            Do While Not rs.EOF
                lsAdditemStr = CStr(.Rows)                                                              '0)Row ��
                lsAdditemStr = lsAdditemStr & vbTab & rs!ZipCode                                        '1)�����ȣ
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!City)                                '2)��.��.��
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Section)                             '3)��.��.��.��
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Village)                             '4)����
                lsAdditemStr = lsAdditemStr & vbTab & Format(rs!Detail1)                                '5)����
                lsAdditemStr = lsAdditemStr & vbTab & Format(rs!Detail2)                                '6)����2
    
                .AddItem lsAdditemStr
    
                rs.MoveNext
            Loop
            rs.Close
    
            '�˻��ڷ� ������ ù�� ����
            If .Rows > .FixedRows Then
                .Row = 1
                .SetFocus
            Else
                MsgBox LoadResString(111), vbInformation
                txtName(1).SetFocus
            End If
        
        End With
    
    End If
    

    Set rs = Nothing
    Set adoCmd = Nothing
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub cmdSelect_Click(Index As Integer)
    Call SelectData(Index)
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub Form_Activate()
    Dim Index As Integer
    
    'S_201312_���_99 �� ���� �߰�
    On Error GoTo Err_Rtn

    
    optSearch(0).Value = True
    
    'S_201312_���_99 �� ���� �߰�
    Index = g_OldNNewClss
    SSTab1.Tab = Index
    
    If Len(txtName(Index)) <> 0 Then
        Call FilllstZip(Index)
    Else
        txtName(Index).SetFocus
    End If
    
    Exit Sub
    
Err_Rtn:
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "frmZip.Form_Activate"
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub Form_Load()

    'S_201312_���_99 �� ���� �߰�
    On Error GoTo Err_Rtn
    
    cmdFind(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdSelect(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdClose(0).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    cmdFind(1).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdSelect(1).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdClose(1).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    'S_201312_���_99 �� ���� �߰�
    Call InitGrid       '�����ȣ �׸��� �ʱ�ȭ

    'S_201312_���_99 �� ���� �߰�
    Call LS_MakeCombo
    
        Exit Sub
    
Err_Rtn:
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "frmZip.Form_Load"
        
End Sub



'S_201312_���_99 �� ���� �߰�
Private Sub InitGrid()

    Call SetVSFlexGrid(grdZipList(0), 3)
   
    With grdZipList(0)
        .Redraw = flexRDNone

        .Rows = 1
        .Cols = 7

        .TextMatrix(0, 0) = "":             .ColWidth(0) = 300
        .TextMatrix(0, 1) = "�����ȣ":     .ColWidth(1) = 800:     .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "���θ��ּ�":   .ColWidth(2) = 4200:     .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "��ȣ":         .ColWidth(3) = 1550:     .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(0, 4) = "":             .ColWidth(4) = 1800:    .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(0, 5) = "":             .ColWidth(5) = 1300:     .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(0, 6) = "":             .ColWidth(6) = 500:      .ColAlignment(6) = flexAlignLeftCenter
        
'        'Ÿ��Ʋ ����
'        .RowHidden(0) = True
        .RowHeight(0) = 30
        
        .ColKey(0) = "IDX":                     .ColKey(1) = UCase("ZipCode"):      .ColKey(2) = UCase("City")
        .ColKey(3) = UCase("Section"):          .ColKey(4) = UCase("Village"):       .ColKey(5) = UCase("Detail1"):
        .ColKey(6) = UCase("Detail2")
        
        .FontSize = 8
        .ExtendLastCol = False          '�������� Ȯ�� ����
        .AllowUserResizing = flexResizeColumns
        
        .Redraw = flexRDDirect
    End With
        
    Call SetVSFlexGrid(grdZipList(1), 3)
   
    With grdZipList(1)
        .Redraw = flexRDNone

        .Rows = 1
        .Cols = 7

        .TextMatrix(0, 0) = "":             .ColWidth(0) = 300
        .TextMatrix(0, 1) = "�����ȣ":     .ColWidth(1) = 800:     .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "��":           .ColWidth(2) = 600:     .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "��.��.��":     .ColWidth(3) = 1550:     .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(0, 4) = "��.��.��.��":  .ColWidth(4) = 1800:    .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(0, 5) = "����":        .ColWidth(5) = 1300:     .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(0, 6) = "":        .ColWidth(6) = 500:      .ColAlignment(6) = flexAlignLeftCenter
        
'        'Ÿ��Ʋ ����
'        .RowHidden(0) = True
        .RowHeight(0) = 30
        
        .ColKey(0) = "IDX":                     .ColKey(1) = UCase("ZipCode"):      .ColKey(2) = UCase("City")
        .ColKey(3) = UCase("Section"):          .ColKey(4) = UCase("Village"):       .ColKey(5) = UCase("Detail1"):
        .ColKey(6) = UCase("Detail2")
        
        .FontSize = 8
        .ExtendLastCol = False          '�������� Ȯ�� ����
        .AllowUserResizing = flexResizeColumns
        
        .Redraw = flexRDDirect
    End With
End Sub

'S_201312_���_99 �� ���� �߰�
Private Sub LS_MakeCombo()
    
    Call LS_MakeComboSiDo
    Call LS_MakeComboSiGunGu

End Sub

'�õ� ComBoBox
'S_201312_���_99 �� ���� �߰�
Private Sub LS_MakeComboSiDo()
    Dim rs As ADODB.Recordset

    cboSiDo.Clear

    If Gf_DB_CM_GetSiDoList(rs, "0") = False Then GoTo Err_Rtn                    '�õ�
    '''    cboSiDo.AddItem "��ü" & Space(50) & "00"  '��ü
    Do While rs.EOF = False
        cboSiDo.AddItem CheckNull(rs!CODE_NAME) & Space(50) & CheckNull(rs!CODE_ID)  '����
        rs.MoveNext
    Loop
    cboSiDo.ListIndex = 0

    
    Exit Sub
    
Err_Rtn:
    Screen.MousePointer = vbDefault
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[ LS_MakeComboSiDo ]"
    End If


End Sub

'�ñ��� ComBoBox
'S_201312_���_99 �� ���� �߰�
Private Sub LS_MakeComboSiGunGu()
    Dim rs As ADODB.Recordset
    
    cboSiGunGu.Clear

    If Gf_DB_CM_GetSiGunGuList(rs, Trim(Mid(cboSiDo.List(cboSiDo.ListIndex), 50)), "0") = False Then GoTo Err_Rtn                    '�ñ���
    cboSiGunGu.AddItem "��ü" & Space(50) & "00000"  '��ü
    Do While rs.EOF = False
        cboSiGunGu.AddItem CheckNull(rs!CODE_NAME) & Space(50) & CheckNull(rs!CODE_ID)  '
        rs.MoveNext
    Loop
    cboSiGunGu.ListIndex = 0
    
        Exit Sub
    
Err_Rtn:
    Screen.MousePointer = vbDefault
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[ LS_MakeComboSiGunGu ]"
    End If


End Sub


'S_201312_���_99 �� ���� �߰�
Private Sub grdZipList_DblClick(Index As Integer)
    If grdZipList(Index).Row <= 0 Then Exit Sub
    
    Call SelectData(Index)
End Sub

'S_201312_���_99 �� ���� �߰�
Private Sub grdZipList_RowColChange(Index As Integer)
    
    If grdZipList(Index).Row <= 0 Then Exit Sub
    
    txtZipCode(Index).Text = grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("ZipCode"))         '�����ȣ

    If Index = 0 Then       '���θ� �ּҸ� ó��
        txtAddr(Index).Text = grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("City"))            '�⺻ �ּ�
        txtAddr2(Index).Text = grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("Section"))           '�ǹ�����+�ι�
        txtAssist(Index).Text = grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("Village"))           '���� �ּ�
        txtGunMoolMngNo(Index).Text = grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("Detail2"))    '�ǹ������ĺ���ȣ
    
    Else                    '�� �����ּ�
        txtAddr(Index).Text = grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("City")) & " " & _
                              grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("Section")) & " " & _
                              grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("Village")) & " "
        txtAddr2(Index).Text = grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("Detail1"))                                                               '�ǹ�����+�ι�
        txtAssist(Index).Text = ""                                                              '���� �ּ�
        txtGunMoolMngNo(Index).Text = ""                                                        '�ǹ������ĺ���ȣ
    End If
    
    
    lstZip(Index).AddItem grdZipList(Index).TextMatrix(grdZipList(Index).Row, grdZipList(Index).ColIndex("ZipCode"))
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub lstZip_DblClick(Index As Integer)
    Call SelectData(Index)
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub lstZip_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call SelectData(Index)
    ElseIf KeyAscii = vbKeyEscape Then
        ReturnStatus = False
        Me.Hide
    End If
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub SelectData(Index)
    If lstZip(Index).ListCount > 0 Then
        ReturnStatus = True
        Me.Hide
    End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    g_OldNNewClss = SSTab1.Tab
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub txtName_GotFocus(Index As Integer)
    With txtName(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'S_201312_���_99 �� ���� ����(OLD:index ������)
Private Sub txtName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FilllstZip(Index)
    ElseIf KeyCode = vbKeyDown Then
        lstZip(Index).SetFocus
    End If
End Sub
