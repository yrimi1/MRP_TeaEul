VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPerson 
   Caption         =   "��� ���� ����(1230)"
   ClientHeight    =   9060
   ClientLeft      =   2070
   ClientTop       =   1350
   ClientWidth     =   11250
   Icon            =   "frmPerson.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   11250
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   6135
      Left            =   30
      TabIndex        =   43
      Top             =   990
      Width           =   3750
      _cx             =   6615
      _cy             =   10821
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
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   1890
      TabIndex        =   42
      Top             =   7170
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      ����(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.OptionButton optSize 
      Caption         =   "��"
      Height          =   330
      Index           =   1
      Left            =   3885
      Style           =   1  '�׷���
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   105
      Value           =   -1  'True
      Width           =   645
   End
   Begin VB.OptionButton optSize 
      Caption         =   "���"
      Height          =   330
      Index           =   0
      Left            =   3885
      Style           =   1  '�׷���
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   495
      Width           =   645
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   0
      Top             =   7500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "����(&S)"
      Height          =   780
      Index           =   3
      Left            =   7185
      MousePointer    =   99  '����� ����
      Style           =   1  '�׷���
      TabIndex        =   25
      ToolTipText     =   "�ڷ� ����"
      Top             =   105
      Visible         =   0   'False
      Width           =   780
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   510
      Left            =   4845
      TabIndex        =   26
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
   Begin VB.CommandButton cmdOperate 
      Caption         =   "�߰�(&A)"
      Height          =   780
      Index           =   0
      Left            =   8775
      MousePointer    =   99  '����� ����
      Style           =   1  '�׷���
      TabIndex        =   29
      ToolTipText     =   "�ڷ� �߰�"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "����(&D)"
      Height          =   780
      Index           =   2
      Left            =   10365
      MousePointer    =   99  '����� ����
      Style           =   1  '�׷���
      TabIndex        =   31
      ToolTipText     =   "�ڷ� ����"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "����(&U)"
      Height          =   780
      Index           =   1
      Left            =   9570
      MousePointer    =   99  '����� ����
      Style           =   1  '�׷���
      TabIndex        =   30
      ToolTipText     =   "�ڷ� ����"
      Top             =   105
      Width           =   780
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "���(&C)"
      Height          =   780
      Index           =   4
      Left            =   7980
      MousePointer    =   99  '����� ����
      Style           =   1  '�׷���
      TabIndex        =   33
      ToolTipText     =   "�ڷ� ���"
      Top             =   105
      Visible         =   0   'False
      Width           =   780
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8010
      TabIndex        =   34
      Top             =   8400
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      �μ�(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   9630
      TabIndex        =   35
      Top             =   8400
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      �ݱ�(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   900
      Left            =   30
      TabIndex        =   36
      Top             =   45
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1588
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboSearch 
         Height          =   300
         Left            =   1380
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   38
         Top             =   120
         Width           =   1770
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   1380
         TabIndex        =   27
         Top             =   495
         Width           =   1755
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   25
         Left            =   60
         TabIndex        =   0
         Top             =   495
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "����� �˻�"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdAll 
         Height          =   330
         Left            =   3195
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   582
         _Version        =   196609
         MousePointer    =   99
         CaptionStyle    =   1
         PictureAnimationEnabled=   0   'False
         Alignment       =   6
         PictureAlignment=   0
         BevelWidth      =   1
         ShapeSize       =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   15
         Left            =   60
         TabIndex        =   37
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "�μ��� �˻�"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   8400
      Left            =   3795
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   30
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   14817
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   741
      TabCaption(0)   =   "  �⺻ ����  "
      TabPicture(0)   =   "frmPerson.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pnlCaption(18)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pnlCaption(16)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraProcess"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "pnlEdit"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pnlMachine"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboTeam"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTemp"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "  �޴� ����  "
      TabPicture(1)   =   "frmPerson.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdMenu"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtTemp 
         Height          =   285
         Left            =   2850
         TabIndex        =   73
         Top             =   8145
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.ComboBox cboTeam 
         Height          =   300
         Left            =   90
         TabIndex        =   24
         Top             =   7140
         Width           =   1215
      End
      Begin Threed.SSPanel pnlMachine 
         Height          =   3735
         Left            =   1335
         TabIndex        =   62
         Top             =   -3195
         Visible         =   0   'False
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   6588
         _Version        =   196609
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSelect 
            Height          =   795
            Left            =   3540
            TabIndex        =   63
            Top             =   2865
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   1402
            _Version        =   196609
            Caption         =   "����"
            PictureAlignment=   9
         End
         Begin VSFlex7LCtl.VSFlexGrid grdMachine 
            Height          =   2325
            Left            =   60
            TabIndex        =   64
            Top             =   480
            Width           =   5220
            _cx             =   9208
            _cy             =   4101
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
         Begin Threed.SSPanel pnlTitle 
            Height          =   420
            Left            =   15
            TabIndex        =   65
            Top             =   15
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   741
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "������"
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdMenu 
         Height          =   6885
         Left            =   -74940
         TabIndex        =   44
         Top             =   930
         Width           =   7290
         _cx             =   12859
         _cy             =   12144
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
      Begin Threed.SSPanel pnlEdit 
         Height          =   5445
         Left            =   90
         TabIndex        =   45
         Top             =   930
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   9604
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSMSYN 
            Caption         =   "���ڹ߼۴��"
            Height          =   255
            Left            =   4140
            TabIndex        =   82
            Top             =   90
            Width           =   1755
         End
         Begin VB.Frame fraAddress 
            Caption         =   "�ּ�"
            Height          =   1965
            Left            =   90
            TabIndex        =   74
            Top             =   2760
            Width           =   7155
            Begin VB.Frame fraDoro 
               Caption         =   "���θ�"
               Height          =   885
               Left            =   1860
               TabIndex        =   79
               Top             =   120
               Width           =   5235
               Begin VB.TextBox txtGunMoolMngNo 
                  Enabled         =   0   'False
                  Height          =   270
                  Left            =   1410
                  TabIndex        =   80
                  TabStop         =   0   'False
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin MRPPlus2.WizText txtAddress1 
                  Height          =   300
                  Left            =   60
                  TabIndex        =   17
                  Top             =   210
                  Width           =   5130
                  _ExtentX        =   9049
                  _ExtentY        =   529
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MRPPlus2.WizText txtAddress2 
                  Height          =   300
                  Left            =   60
                  TabIndex        =   18
                  Top             =   540
                  Width           =   3390
                  _ExtentX        =   5980
                  _ExtentY        =   529
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MRPPlus2.WizText txtAddressAssist 
                  Height          =   300
                  Left            =   3450
                  TabIndex        =   19
                  Top             =   540
                  Width           =   1740
                  _ExtentX        =   3069
                  _ExtentY        =   529
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin VB.Frame fraOldNNew 
               Height          =   405
               Left            =   60
               TabIndex        =   78
               Top             =   180
               Width           =   1785
               Begin VB.OptionButton optOldNNew 
                  Caption         =   "���θ�"
                  Height          =   225
                  Index           =   0
                  Left            =   60
                  TabIndex        =   15
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton optOldNNew 
                  Caption         =   "����"
                  Height          =   225
                  Index           =   1
                  Left            =   1020
                  TabIndex        =   16
                  Top             =   120
                  Width           =   675
               End
            End
            Begin VB.Frame fraJiBun 
               Caption         =   "����"
               Height          =   885
               Left            =   1860
               TabIndex        =   77
               Top             =   1050
               Width           =   5235
               Begin MRPPlus2.WizText txtAddress 
                  Height          =   300
                  Index           =   0
                  Left            =   60
                  TabIndex        =   20
                  Top             =   180
                  Width           =   5130
                  _ExtentX        =   9049
                  _ExtentY        =   529
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MRPPlus2.WizText txtAddress 
                  Height          =   300
                  Index           =   1
                  Left            =   60
                  TabIndex        =   21
                  Top             =   495
                  Width           =   5115
                  _ExtentX        =   9022
                  _ExtentY        =   529
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin MSMask.MaskEdBox mskZipCode 
               Height          =   300
               Left            =   60
               TabIndex        =   75
               Top             =   630
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "###-###"
               PromptChar      =   "_"
            End
            Begin Threed.SSCommand cmdFind 
               Height          =   300
               Left            =   870
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   630
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   529
               _Version        =   196609
               ButtonStyle     =   3
               Outline         =   0   'False
            End
         End
         Begin VB.TextBox txtEMail 
            Height          =   300
            Left            =   1350
            TabIndex        =   22
            Top             =   4755
            Width           =   5805
         End
         Begin VB.ComboBox cboDuty 
            Height          =   300
            ItemData        =   "frmPerson.frx":0044
            Left            =   1365
            List            =   "frmPerson.frx":0046
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   4
            Top             =   1500
            Width           =   1740
         End
         Begin VB.ComboBox cboDepart 
            Height          =   300
            Left            =   1365
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   3
            Top             =   1170
            Width           =   1740
         End
         Begin VB.ComboBox cboSolarClss 
            Height          =   300
            Left            =   2985
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   12
            Top             =   2430
            Width           =   705
         End
         Begin VB.Frame fraName 
            Height          =   870
            Left            =   4095
            TabIndex        =   47
            Top             =   420
            Width           =   3135
            Begin VB.TextBox txtPassWord 
               BackColor       =   &H00FFC0C0&
               Height          =   300
               IMEMode         =   3  '��� ����
               Left            =   1340
               PasswordChar    =   "*"
               TabIndex        =   6
               Top             =   495
               Width           =   1665
            End
            Begin VB.TextBox txtUserID 
               BackColor       =   &H00FFC0C0&
               Height          =   300
               Left            =   1340
               MaxLength       =   15
               TabIndex        =   5
               Top             =   165
               Width           =   1665
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   2
               Left            =   105
               TabIndex        =   48
               Top             =   165
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "��  ��  ��"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   14
               Left            =   105
               TabIndex        =   49
               Top             =   495
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "��� ��ȣ"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin VB.TextBox txtRemark 
            Height          =   330
            Left            =   1350
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   23
            Top             =   5085
            Width           =   5850
         End
         Begin MRPPlus2.WizText txtTelePhone 
            Height          =   300
            Left            =   5400
            TabIndex        =   14
            Top             =   2430
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MRPPlus2.WizText txtHandPhone 
            Height          =   300
            Left            =   5400
            TabIndex        =   13
            Top             =   2100
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MRPPlus2.WizText txtName 
            Height          =   300
            Left            =   1365
            TabIndex        =   1
            Top             =   450
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MRPPlus2.WizText txtCode 
            Height          =   300
            Left            =   1365
            TabIndex        =   46
            Top             =   90
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12648384
         End
         Begin MSMask.MaskEdBox mskStartDate 
            Height          =   300
            Left            =   5415
            TabIndex        =   7
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####�� ##�� ##��"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskRegistID 
            Height          =   300
            Left            =   1350
            TabIndex        =   10
            Top             =   2100
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "######-#######"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   8
            Left            =   4155
            TabIndex        =   50
            Top             =   2430
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��ȭ ��ȣ"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   105
            TabIndex        =   51
            Top             =   450
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��    ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   5
            Left            =   4095
            TabIndex        =   52
            Top             =   1320
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "�Ի� ����"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   6
            Left            =   105
            TabIndex        =   53
            Top             =   2100
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "�ֹε�Ϲ�ȣ"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   4
            Left            =   105
            TabIndex        =   54
            Top             =   1170
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��    ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   7
            Left            =   4155
            TabIndex        =   55
            Top             =   2100
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��  ��  ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   9
            Left            =   105
            TabIndex        =   56
            Top             =   2430
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "�������"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   12
            Left            =   105
            TabIndex        =   57
            Top             =   5100
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��    ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   13
            Left            =   105
            TabIndex        =   58
            Top             =   1500
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��    å"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSMask.MaskEdBox mskBirthday 
            Height          =   300
            Left            =   1350
            TabIndex        =   11
            Top             =   2430
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####�� ##�� ##��"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   3
            Left            =   4110
            TabIndex        =   59
            Top             =   1635
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   196609
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkEnd 
               Caption         =   "�������"
               Height          =   255
               Left            =   90
               TabIndex        =   8
               Top             =   30
               Width           =   1035
            End
         End
         Begin MSMask.MaskEdBox mskEndDate 
            Height          =   300
            Left            =   5415
            TabIndex        =   9
            Top             =   1635
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####�� ##�� ##��"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   17
            Left            =   105
            TabIndex        =   60
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��    ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   0
            Left            =   105
            TabIndex        =   61
            Top             =   4770
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "E-Mail"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MRPPlus2.WizText txtename 
            Height          =   300
            Left            =   1365
            TabIndex        =   2
            Top             =   810
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   10
            Left            =   105
            TabIndex        =   81
            Top             =   810
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "�����̸�"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   0
            X2              =   7455
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000003&
            Index           =   1
            X1              =   0
            X2              =   7470
            Y1              =   1995
            Y2              =   1995
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   15
            X2              =   7470
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000003&
            Index           =   3
            X1              =   0
            X2              =   7455
            Y1              =   405
            Y2              =   405
         End
      End
      Begin Threed.SSFrame fraProcess 
         Height          =   1470
         Left            =   1335
         TabIndex        =   66
         Top             =   6405
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2593
         _Version        =   196609
         Enabled         =   0   'False
         Begin Threed.SSCommand cmdMachine 
            Height          =   390
            Index           =   1
            Left            =   5070
            TabIndex        =   67
            Top             =   495
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   688
            _Version        =   196609
            Caption         =   "����"
         End
         Begin Threed.SSCommand cmdMachine 
            Height          =   390
            Index           =   0
            Left            =   5070
            TabIndex        =   68
            Top             =   60
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   688
            _Version        =   196609
            Caption         =   "�߰�"
         End
         Begin VSFlex7LCtl.VSFlexGrid grdProcess 
            Height          =   1335
            Left            =   60
            TabIndex        =   69
            Top             =   75
            Width           =   4965
            _cx             =   8758
            _cy             =   2355
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
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   16
         Left            =   90
         TabIndex        =   70
         Top             =   6435
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "�۾�����"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   18
         Left            =   90
         TabIndex        =   71
         Top             =   6795
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "�� �� ��"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.Label lblBoard 
      Caption         =   "�� ��� �⺻ ���� �Է� �� ��� �޴��� ���� �Ͻʽÿ�."
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   3585
      TabIndex        =   72
      Top             =   8760
      Width           =   4395
   End
   Begin VB.Image imgUnCheck 
      Height          =   165
      Left            =   3030
      Picture         =   "frmPerson.frx":0048
      Top             =   6195
      Width           =   165
   End
   Begin VB.Image ImgCheck 
      Height          =   165
      Left            =   3030
      Picture         =   "frmPerson.frx":0120
      Top             =   5970
      Width           =   165
   End
   Begin VB.Image imgItem 
      Height          =   195
      Left            =   3240
      Picture         =   "frmPerson.frx":01F8
      Top             =   6000
      Width           =   195
   End
   Begin VB.Image imgFolder 
      Height          =   195
      Left            =   3240
      Picture         =   "frmPerson.frx":0320
      Top             =   6225
      Width           =   195
   End
   Begin VB.Label lblCount 
      Caption         =   "�˻��Ǽ� :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   30
      TabIndex        =   41
      Top             =   7260
      Width           =   3630
   End
End
Attribute VB_Name = "frmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'** System �� : MRRPLUS2
'** Author    : Wizard
'** �ۼ���    :
'** ����      : �ŷ�ó ���
'** ��������  :
'** ��������  : 2013.11.25
'**------------------------------------------------------------------------------------------------
'
'  ��û���� ID: S_201312_��������_99
'  ��û��:
'  ���泯¥ : 2013.11.25
'  �۾���   : ���¿�
'  ��û���� : �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�
'  ���泻�� : ���θ�,�� �����ּ� �ɼ� ��ư �߰�
'**************************************************************************************************

Option Explicit

Private m_sFlag As String * 1

Private Const REPORTFILE = "\Report\Person.rpt"

Private Const LIMIT_ROW1 = 20
Private Const LIMIT_ROW2 = 19
Private Const LIMIT_WIDTH1 = 1400
Private Const LIMIT_WIDTH2 = 3930

Private m_bloading As Boolean



Private Sub cmdMachine_Click(Index As Integer)
    
    If Index = 0 Then
        With grdProcess
            .Rows = .Rows + 1
    
            .Cell(flexcpPicture, .Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, .Rows - 1, 2) = flexPicAlignCenterCenter
            
            .Cell(flexcpPicture, .Rows - 1, 5) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, .Rows - 1, 5) = flexPicAlignCenterCenter
            
            .SetFocus
            .Select .Rows - 1, 1
        End With
    Else
        With grdProcess
            If .Rows = .FixedRows Or .Row < .FixedRows Or .Row >= .Rows Then Exit Sub

            .RemoveItem .Row
    
        End With
    End If
End Sub

Private Sub cmdSelect_Click()
    Dim sMachine$, sMachineID$, sMachineNO$
    
    With grdMachine
        If .Row = 0 Then Exit Sub
        
        sMachine = .TextMatrix(.Row, 1)
        sMachineNO = .TextMatrix(.Row, 2)
        sMachineID = .TextMatrix(.Row, 3)
    End With
    pnlMachine.Visible = False
    
    With grdProcess
        .TextMatrix(.Row, 3) = sMachine
        .TextMatrix(.Row, 4) = sMachineNO
        .TextMatrix(.Row, 7) = sMachineID
        
    End With
End Sub

Private Sub Form_Load()
    'S_201312_��������_99 �� ���� ����
''    Me.Move 0, 0, 11355, 8325
    Me.Move 0, 0, 11355, 9630

    On Error GoTo ErrHandler

    Call SetOperate(Me)

    With cboSolarClss
        .AddItem "���"
        .AddItem "����"

        .ListIndex = 0
    End With

    cmdFind.Picture = LoadResPicture("FIND", vbResIcon)
    cmdAll.Picture = LoadResPicture("ALL", vbResIcon)

    lblCount.Caption = LoadResString(250)

    m_bloading = False

    Call InitGrid
    Call MakeCodeCombo(cboTeam, CD_TEAM)
    Call MakeMenu

    MousePointer = vbHourglass

    Call MakeCodeCombo(cboSearch, CD_DEPART, True)
    Call MakeCodeCombo(cboDepart, CD_DEPART)
    Call MakeCodeCombo(cboDuty, CD_DUTY)
    
    grdMenu.Editable = flexEDNone
    
    MousePointer = vbDefault
    
    tabMain.Tab = 0
    
    Exit Sub

ErrHandler:
    MousePointer = vbDefault
    
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Sub

Private Sub cboSearch_Click()
    On Error Resume Next
   
    With cboSearch
        If .ListIndex = 0 Then
            Call FillGrid
        Else
            Call FillGrid(Format(.ItemData(.ListIndex), "00"))
        End If
    End With
End Sub

Private Sub chkEnd_Click()
    If chkEnd = vbChecked Then
        mskEndDate.Enabled = True
    Else
        mskEndDate.Enabled = False
        mskEndDate.Text = ""
    End If
End Sub

Private Sub grdData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdOperate_Click(ID_UPDATE)
    End If
End Sub



Private Sub grdMachine_DblClick()
    Call cmdSelect_Click
End Sub


Private Sub grdMenu_Click()
    Dim nMenuID%, nSubMenu%
    Dim iLoop%, colSeq%
    Dim sParentID As String
    Dim Checked As Boolean
    
    On Error GoTo ErrHandler:
    
    If Not cmdOperate(3).Visible = True Then
        Exit Sub
    End If
    
    With grdMenu
        If .Col < 2 Then Exit Sub
        
        nMenuID = val(Left(Right(.TextMatrix(.Row, 1), 5), 4))  '���õ� Row�� �޴�ID
        If nMenuID Mod 100 = 0 Then
            Checked = IIf(.Cell(flexcpPicture, .Row, .Col) = ImgCheck, False, True)
            .Cell(flexcpPicture, .Row, .Col) = IIf(Checked, ImgCheck, imgUnCheck)
        Else
            Checked = IIf(.Cell(flexcpChecked, .Row, .Col) = flexChecked, False, True)  'üũ�Ǹ� true, üũ������ false
            .Cell(flexcpChecked, .Row, .Col) = Checked
        End If
        
        
        .Redraw = flexRDNone
        
        If .Col = 2 Then '��뱸��.
            If .RowOutlineLevel(.Row) = 0 Then  '�ֻ��� ��� üũ. ��ü ������� ����
                For iLoop = .FixedRows To .Rows - 1
                    For colSeq = 2 To 6
                        If val(Left(Right(.TextMatrix(iLoop, 1), 5), 4)) Mod 100 = 0 Then
                            .Cell(flexcpPicture, iLoop, colSeq) = IIf(Checked, ImgCheck, imgUnCheck)
                        Else
                            .Cell(flexcpChecked, iLoop, colSeq) = Checked
                        End If
                        
                    Next colSeq
                Next iLoop
                
            Else  ' �ֻ����� �ƴ� ����..
            
                For colSeq = 2 To 6  '���õ� Row�� üũ�ڽ� üũ
                    If nMenuID Mod 100 = 0 Then
                        .Cell(flexcpPicture, .Row, colSeq) = IIf(Checked, ImgCheck, imgUnCheck)
                    Else
                        .Cell(flexcpChecked, .Row, colSeq) = Checked
                    End If
                Next colSeq
                  
                If Not .RowOutlineLevel(.Row) = 4 Then  '������ ��� �ƴ϶��...
                    nSubMenu = .Row + 1  '���õ� ����� ���� ������ üũ..
                    Do While .RowOutlineLevel(nSubMenu) > .RowOutlineLevel(.Row)
                        For colSeq = 2 To 6
                            If val(Left(Right(.TextMatrix(nSubMenu, 1), 5), 4)) Mod 100 = 0 Then
                                .Cell(flexcpPicture, nSubMenu, colSeq) = IIf(Checked, ImgCheck, imgUnCheck)
                            Else
                                .Cell(flexcpChecked, nSubMenu, colSeq) = Checked
                            End If
                            
                        Next colSeq
                        
                        nSubMenu = nSubMenu + 1
                        If nSubMenu > .Rows - 1 Then
                            Exit Do
                        End If
                    Loop
                 End If
                    
                    ' ���� ��忡 üũ�ϱ�..
                 If Checked Then
                    For colSeq = 2 To 6
                        .Cell(flexcpChecked, 1, colSeq) = True
                    Next colSeq
                    
                    If Not .TextMatrix(.Row, 7) = "" Then
                        sParentID = .TextMatrix(.Row, 7)
                        
                        Do While Not sParentID = ""
                    
                            sParentID = UpperNodeCheck(sParentID)
                        Loop
                    End If
                        
                 End If
           
            End If
        
        Else  '�߰�, ����, ����, ����
            If .RowOutlineLevel(.Row) = 0 Then  '�ֻ��� ��� üũ. ��ü ������� ����
                For iLoop = .FixedRows To .Rows - 1
                    If val(Left(Right(.TextMatrix(iLoop, 1), 5), 4)) Mod 100 = 0 Then
                        .Cell(flexcpPicture, iLoop, .Col) = IIf(Checked, ImgCheck, imgUnCheck)
                        If Checked Then
                            .Cell(flexcpPicture, iLoop, 2) = ImgCheck
                        End If
                    Else
                        .Cell(flexcpChecked, iLoop, .Col) = Checked
                        If Checked Then
                            .Cell(flexcpChecked, iLoop, 2) = Checked
                        End If
                    End If
              
                Next iLoop
            Else
                
                  '���õ� Row�� üũ�ڽ� üũ
                .Cell(flexcpChecked, .Row, .Col) = Checked
                If Checked Then   '���õ� �޴��� '��뱸��' üũ
                    If nMenuID Mod 100 = 0 Then
                        .Cell(flexcpPicture, .Row, 2) = IIf(Checked, ImgCheck, imgUnCheck)
                    Else
                        .Cell(flexcpChecked, .Row, 2) = Checked
                    End If
                    
                    If Not .TextMatrix(.Row, 7) = "" Then
                        sParentID = .TextMatrix(.Row, 7)
                        
                        Do While Not sParentID = ""  '�߰�, ������ �Ϻ� �÷��� ���ý�..
                            sParentID = UpperNodeCheck(sParentID, False)
                        Loop
                    End If
                    
                    
                End If
                
                  
                If Not .RowOutlineLevel(.Row) = 4 Then  '������ ��� �ƴ϶��...
                    nSubMenu = .Row + 1  '���õ� ����� ���� ������ üũ..
                    Do While .RowOutlineLevel(nSubMenu) > .RowOutlineLevel(.Row)
                        If val(Left(Right(.TextMatrix(nSubMenu, 1), 5), 4)) Mod 100 = 0 Then
                            .Cell(flexcpPicture, nSubMenu, .Col) = IIf(Checked, ImgCheck, imgUnCheck)
                            If Checked Then
                                .Cell(flexcpPicture, nSubMenu, 2) = ImgCheck
                            End If
                        Else
                            .Cell(flexcpChecked, nSubMenu, .Col) = Checked
                            If Checked Then
                                .Cell(flexcpChecked, nSubMenu, 2) = Checked
                            End If
                        End If
                        
                        nSubMenu = nSubMenu + 1
                        If nSubMenu > .Rows - 1 Then
                            Exit Do
                        End If
                    Loop
                 End If
             End If
                    
        End If
        
        .Redraw = flexRDDirect
        
    End With
Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
    
End Sub
Function UpperNodeCheck(sParentID As String, Optional Mode As Boolean)
    Dim iLoop%, irow%
    Dim colSeq%

    With grdMenu
    
        For iLoop = .FixedRows To .Rows - 1
            If val(Left(Right(.TextMatrix(iLoop, 1), 5), 4)) = val(sParentID) Then
                If Mode = True Then  ' ��ü �÷� üũ

                    For colSeq = 2 To 6

                        .Cell(flexcpPicture, iLoop, colSeq) = ImgCheck
                    Next colSeq
                Else   ' ��뱸�� �÷��� üũ
                    .Cell(flexcpPicture, iLoop, 2) = ImgCheck
                End If

                sParentID = .TextMatrix(iLoop, 7)
                UpperNodeCheck = sParentID
                Exit Function

            End If

        Next iLoop
        
    End With
    UpperNodeCheck = ""


End Function


Private Sub grdProcess_Click()
    Dim sProcID$
    
    With grdProcess
    
        If .Rows = .FixedRows Then Exit Sub
        
        If .Col = 2 Then
            sProcID = txtTemp
            txtTemp = ""
            If ReturnCode(LG_PROCESS, , False, txtTemp) = True Then
                
                .TextMatrix(.Row, 1) = txtTemp
                .TextMatrix(.Row, 6) = txtTemp.Tag
            End If
            
            If .TextMatrix(.Row, 6) <> sProcID Then
                .TextMatrix(.Row, 3) = ""
                .TextMatrix(.Row, 4) = ""
                .TextMatrix(.Row, 7) = ""
                
            End If
            
        ElseIf .Col = 5 Then
            txtTemp = .TextMatrix(.Row, 6)
            If Len(Trim(txtTemp)) = 0 Then
                MsgBox "���� ������ ������ �ֽʽÿ�"
                Exit Sub
            End If
            
            pnlTitle = "  " & .TextMatrix(.Row, 1) & " ���� ����"
            pnlMachine.Visible = True
            Call FillGridMachine
            
        End If
        
    End With
End Sub



Private Sub FillGridMachine()
    Dim oProcess As PlusLib2.CProcess
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    Set oProcess = New PlusLib2.CProcess
    oProcess.Connection = g_adoCon
    
    Set rs = oProcess.GetMachine(txtTemp)
    Set oProcess = Nothing
    
    With grdMachine
        .Rows = .FixedRows
        
        Do Until rs.EOF
            i = i + 1
            .AddItem CStr(i) & vbTab & rs!Machine & vbTab & rs!MachineNO & vbTab & rs!machineid
        
            rs.MoveNext
        Loop
        
        rs.Close
        Set rs = Nothing
    
    End With
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oProcess = Nothing
    
    Call ErrorBox(Err.Number, Err.Source, Err.Description)

End Sub



'S_201312_��������_99 �� ���� �߰�
Private Sub optOldNNew_Click(Index As Integer)
    If optOldNNew(0).Value = True Then
        fraDoro.Enabled = True
        fraJiBun.Enabled = False
    Else
        fraDoro.Enabled = False
        fraJiBun.Enabled = True
    End If
End Sub

'S_201312_��������_99 �� ���� �߰�
Private Sub txtAddress1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Call cmdFind_Click
    End If
    
End Sub


Private Sub txtSearch_Change()
    Dim i%, iCount%, iNowRow%

    On Error GoTo ErrHandler

    If Len(Trim(txtSearch)) > 0 Then
        With grdData
            .Redraw = flexRDNone

            For i = .FixedRows To .Rows - .FixedRows
                If InStr(UCase(.TextArray(i * .Cols + 2)), UCase(txtSearch)) = 0 Then
                    .RowHidden(i) = True
                    iCount = iCount + 1
                Else
                    .RowHidden(i) = False
                    iNowRow = i
                End If
            Next i

            If iNowRow > .FixedRows Then
                .Row = iNowRow

                .Col = .FixedCols
                .ColSel = .Cols - 1
            End If

            .Redraw = flexRDDirect

            .TopRow = .Row
        End With
    Else
        Call cmdAll_Click
    End If

    cmdAll.Visible = IIf(iCount > 0, True, False)

    Call ChangeScroll

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "Person.txtSearch_Change", Err.Description)
End Sub

Private Sub cmdAll_Click()
    Dim i%

    With grdData
        .Redraw = flexRDNone

        For i = .FixedRows To .Rows - .FixedRows
            .RowHidden(i) = False
        Next i

        .Redraw = flexRDDirect
    End With

    txtSearch.Text = ""
    cmdAll.Visible = False
End Sub

Private Sub grdData_RowColChange()
    Call ShowData
End Sub

Private Sub grdData_DblClick()
    With grdData
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub

        Call cmdOperate_Click(ID_UPDATE)
    End With
End Sub

Private Sub optSize_Click(Index As Integer)
    If optSize(0).Value Then    'Ȯ��
        grdData.Width = 11235
        tabMain.Visible = False
    Else                        '���
        grdData.Width = 3750
        tabMain.Visible = True
    End If
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim nMaxValue As String

    On Error GoTo ErrHandler

    If optSize(0).Value Then optSize(1).Value = True

    Select Case Index
    Case ID_ADDNEW
        m_sFlag = ID_ADDNEW
        Call ChangeMode(Me, False)
        
        'S_201312_��������_99 �� ���� �߰�-----------------------------------------------
        If optOldNNew(0).Value = True Then
            fraDoro.Enabled = True
            fraJiBun.Enabled = False
        Else
            fraDoro.Enabled = False
            fraJiBun.Enabled = True
        End If
        '-------------------------------------------------------------------------


        If tabMain.Tab = 0 Then
            Call ClearData
            txtUserID.Locked = False
            txtCode.Locked = False
            txtName.SetFocus
        End If
        Call MakeMenu
        fraProcess.Enabled = True
       
        pnlMsg.Caption = LoadResString(302)
    Case ID_UPDATE
        m_sFlag = ID_UPDATE
        Call ChangeMode(Me, False)
        
        'S_201312_��������_99 �� ���� �߰�-----------------------------------------------
        If optOldNNew(0).Value = True Then
            fraDoro.Enabled = True
            fraJiBun.Enabled = False
        Else
            fraDoro.Enabled = False
            fraJiBun.Enabled = True
        End If
        '-------------------------------------------------------------------------
        
        txtUserID.Locked = True
        txtCode.Locked = True
        
        fraProcess.Enabled = True
       
        txtName.SetFocus
        
        pnlMsg.Caption = LoadResString(303)
    Case ID_DELETE
        If grdData.Rows = grdData.FixedRows Then Exit Sub

        If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "����Ȯ��") = vbYes Then
            If DeleteData() Then
                With cboSearch
                    If .ListIndex = 0 Then
                        Call FillGrid
                    Else
                        Call FillGrid(Format(.ItemData(.ListIndex), "00"))
                    End If
                End With
            End If
        End If
    Case ID_SAVE
        If Not CheckData() Then Exit Sub

        If SaveData() Then
            grdMenu.Editable = flexEDNone
            Call ChangeMode(Me, True)
            With cboSearch
                If .ListIndex = 0 Then
                    Call FillGrid
                Else
                    Call FillGrid(Format(.ItemData(.ListIndex), "00"))
                End If
            End With
            fraProcess.Enabled = False
       
            txtUserID.Locked = False
        End If
        grdData.SetFocus

    Case ID_CANCEL
        txtUserID.Locked = False
        Call ChangeMode(Me, True)
        
        fraProcess.Enabled = False
       
        With cboSearch
            If .ListIndex = 0 Then
                Call FillGrid
            Else
                Call FillGrid(Format(.ItemData(.ListIndex), "00"))
            End If
        End With
        grdData.SetFocus
    End Select

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "Person.cmdOperate_Click", Err.Description)
End Sub


Private Sub cboDepart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub cboDuty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub



Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        grdData.SetFocus
    End If
End Sub

Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call NextFocus
End Sub

Private Sub txtPassword_GotFocus()
    Call GotFocusText(txtPassWord)
End Sub

Private Sub txtPassWord_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
    
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call NextFocus
End Sub

Private Sub mskStartDate_GotFocus()
    mskStartDate.SelStart = 0
    mskStartDate.SelLength = 13
End Sub

Private Sub mskStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub mskStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call NextFocus
End Sub

Private Sub mskEndDate_GotFocus()
    With mskStartDate
        .SelStart = 0
        .SelLength = 13
    End With
End Sub

Private Sub mskEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub mskEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call NextFocus
End Sub

Private Sub mskRegistID_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub mskRegistID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call NextFocus
End Sub

Private Sub mskRegistID_Validate(Cancel As Boolean)
    If Len(mskRegistID) >= 6 Then
        If CInt(Mid(mskRegistID, 1, 2)) > CInt(Mid(MakeDate(DF_SHORT, Now), 3, 2)) Then
            mskBirthday = "19" & Left(mskRegistID, 6)
        Else
            mskBirthday = "20" & Left(mskRegistID, 6)
        End If
    End If
End Sub

Private Sub cboSolarClss_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub mskBirthday_GotFocus()
    mskBirthday.SelStart = 0
    mskBirthday.SelLength = 13
End Sub

Private Sub mskBirthday_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub mskBirthday_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call NextFocus
End Sub



Private Sub txtAddress_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            Call cmdFind_Click
        End If
    End If
End Sub

'''S_201312_��������_99 �� ���� ����-OLD�ҽ�
''Private Sub cmdFind_Click()
''    Dim oZipFind As PlusFind2.CZipFind
''
''    Set oZipFind = New PlusFind2.CZipFind
''    oZipFind.Connection = g_adoCon
''    oZipFind.Address1 = txtAddress(0)
''
''    If oZipFind.Show() Then
''        txtAddress(0) = oZipFind.Address
''        mskZipCode = oZipFind.ZipCode
''    End If
''    Set oZipFind = Nothing
''
''    txtAddress(1).SetFocus
''End Sub

'S_201312_��������_99 �� ���� ����-NEW�ҽ�
Private Sub cmdFind_Click()
    Dim oZipFind As PlusFind2.CZipFind

    On Error GoTo ErrHandler
    
    'S_201312_��������_99 �� ���� �߰�
    '������ �����ȣ  DB ���� �����
    If g_bChkWizDBConn = False Then
        g_bChkWizDBConn = PlusMDI.ConnectWizDB()
    End If
    
    
    Set oZipFind = New PlusFind2.CZipFind
    'S_201312_��������_99 �� ���� ����(OLD: g_adoCon)
    'oZipFind.DBGubun = g_sDBGubun        'S_201102_â���_01 �� ���� �߰�
    oZipFind.Connection = g_adoWizCon

    'S_201312_��������_99 �� ���� �߰�
    If optOldNNew(0).Value = True Then      '���θ� �ּ�
        oZipFind.Address1 = txtAddress1
    Else                                    '���� �ּ�
        'S_201312_��������_99 �� ���� ����(OLD:oZipFind.Address1)
        oZipFind.AddressJiBun1 = txtAddress(0).Text
    End If

    'S_201312_��������_99 �� ���� �߰�
    oZipFind.OldNNewSet = IIf(optOldNNew(0).Value = True, "0", "1")
    
    If oZipFind.Show() Then
        mskZipCode = oZipFind.ZipCode
        
        'S_201312_��������_99 �� ���� ����-----------------------------------------------
''        txtAddress(0) = oZipFind.Address
        If oZipFind.OldNNewClss = "0" Then    '���θ� �ּ�
            optOldNNew(0).Value = True
                
            txtAddress1.Text = oZipFind.Address
            txtAddress2.Text = oZipFind.AddressDetail
            txtAddressAssist.Text = oZipFind.AddressAssist
            txtGunMoolMngNo.Text = oZipFind.GunMoolMngNo

            txtAddress2.SetFocus
        Else
            optOldNNew(1).Value = True
            txtAddress(0).Text = oZipFind.Address
            txtAddress(1).Text = ""                       'S_201312_��������_99 �� ���� �߰�
        
            txtAddress(1).SetFocus
        End If
        '----------------------------------------------------------------------------
        
    End If
    Set oZipFind = Nothing

'''    txtAddress(1).SetFocus
    Exit Sub
ErrHandler:
    Set oZipFind = Nothing
    
    Call ErrorBox(Err.Number, "frmPerson.cmdFind_Click", Err.Description)
    
End Sub

Private Sub mskZipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub mskZipCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call NextFocus
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then Call NextFocus
' �������� ������ �Է��� �� �־�� ���� ������?
End Sub

Private Sub grdMenu_DblClick()
    With grdMenu
        If .MouseCol <> 1 Or .MouseRow < 1 Then Exit Sub

        .IsCollapsed(.Row) = IIf(.IsCollapsed(.Row) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
    End With
End Sub

Private Sub cmdExcel_Click()
    If grdData.Rows = 1 Then
        Call MessageBox(LoadResString(111))
        Exit Sub
    End If

    Call MakeExcelGrid(grdData)
End Sub

Private Sub cmdPrint_Click()
    Dim oPerson As PlusLib2.CPerson
    Dim rs As ADODB.Recordset
    Dim sParam() As String
    Dim nOut%
    Dim sDepart$, sName$

    On Error GoTo ErrHandler
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    ' Printing
    Screen.MousePointer = vbHourglass
    
    Set oPerson = New PlusLib2.CPerson
    oPerson.Connection = g_adoCon
    sDepart = IIf(cboSearch.ListIndex = 0, "", "0" & cboSearch.ItemData(cboSearch.ListIndex))
    sName = IIf(Len(txtSearch) > 0, CheckNull(txtSearch), "%")
    Set rs = oPerson.GetPerson(sDepart)
    Set oPerson = Nothing
    
    ReDim sParam(2)
    sParam(0) = "�ŷ�ó ����Ʈ"
    sParam(1) = CompanyName
    sParam(2) = "�μ��� : " & cboSearch
 '   sParam(3) = "��   �� : " & IIf(Len(txtSearch) > 0, txtSearch, "(��ü)")
    
    If PlusMDI.PrintPreview Then
        nOut = 0
    Else
        nOut = 1
    End If
    
    Call PrintReport(REPORTFILE, rs, sParam, PlusMDI.mnuPopup)
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, "frmPerson.cmdPrint_Click", Err.Description)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    With grdData
        .Cols = 29             'S_201312_��������_99 �� ���� ���� (OLD:22)
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .Rows = 1

        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "":                .ColWidth(1) = 250:             .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "����":            .ColWidth(2) = 960:             .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "���ID":          .ColWidth(3) = 900:             .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(0, 4) = "����ڹ�ȣ":      .ColWidth(4) = LIMIT_WIDTH1:    .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(0, 5) = "�μ�":            .ColWidth(5) = 1440:            .ColAlignment(5) = flexAlignLeftCenter
        .TextMatrix(0, 6) = "��å":            .ColWidth(6) = 1065:            .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(0, 7) = "�Ի�����":        .ColWidth(7) = 1000:            .ColAlignment(7) = flexAlignCenterCenter
        .TextMatrix(0, 8) = "�ֹε�Ϲ�ȣ":    .ColWidth(8) = 1400:            .ColAlignment(8) = flexAlignCenterCenter
        .TextMatrix(0, 9) = "��ȭ��ȣ":        .ColWidth(9) = 1300:            .ColAlignment(9) = flexAlignLeftCenter
        .TextMatrix(0, 10) = "�������":       .ColWidth(10) = 1250:           .ColAlignment(10) = flexAlignLeftCenter
        .TextMatrix(0, 11) = "�����ּ�1":        .ColWidth(12) = 0
        .TextMatrix(0, 12) = "�����ּ�2":        .ColWidth(13) = 0
        .TextMatrix(0, 13) = "�����ȣ":       .ColWidth(14) = 0
        .TextMatrix(0, 14) = "������":       .ColWidth(15) = 0
        .TextMatrix(0, 15) = "�������":       .ColWidth(16) = 0
        .TextMatrix(0, 16) = "Password":       .ColWidth(17) = 0
        .TextMatrix(0, 17) = "DepartID":       .ColWidth(18) = 0
        .TextMatrix(0, 18) = "DutyID":         .ColWidth(19) = 0
        .TextMatrix(0, 19) = "�ڵ���":         .ColWidth(20) = 0
        .TextMatrix(0, 20) = "EMail":          .ColWidth(21) = 0
        .TextMatrix(0, 21) = "�۾���":          .ColWidth(21) = 0
        .TextMatrix(0, 22) = "�����̸�":                .ColWidth(22) = 0
        'S_201312_��������_99 �� ���� �߰�-------------------------------------
        .TextMatrix(0, 23) = "���ڹ߼۴�󿩺�":  .ColWidth(23) = 0
        .TextMatrix(0, 24) = "�ּұ���":          .ColWidth(24) = 0
        .TextMatrix(0, 25) = "�ǹ��ĺ���ȣ":          .ColWidth(25) = 0
        .TextMatrix(0, 26) = "���θ��ּ�1":          .ColWidth(26) = 0
        .TextMatrix(0, 27) = "���θ��ּ�2":          .ColWidth(27) = 0
        .TextMatrix(0, 28) = "���θ� �����ּ�":         .ColWidth(28) = 0
        
        '//�� ����ColKey ����
        .ColKey(0) = "Idx"
        .ColKey(1) = "DspEnd"
        .ColKey(2) = "Name"
        .ColKey(3) = "PersonID"
        .ColKey(4) = "UserID"
        .ColKey(5) = "Depart"
        .ColKey(6) = "Duty"
        .ColKey(7) = "StartDate"
        .ColKey(8) = "RegistID"
        .ColKey(9) = "Phone"
        .ColKey(10) = "BirthDay"
        'S_201312_��������_99 �� ���� ����(OLD:Address1)
        .ColKey(11) = "AddressJiBun1"
        'S_201312_��������_99 �� ���� ����(OLD:Address2)
        .ColKey(12) = "AddressJiBun2"
        .ColKey(13) = "ZipCode"
        .ColKey(14) = "Remark"
        .ColKey(15) = "EndDate"
        .ColKey(16) = "Password"
        .ColKey(17) = "DepartID"
        .ColKey(18) = "DutyID"
        .ColKey(19) = "HandPhone"
        .ColKey(20) = "Email"
        .ColKey(21) = "TeamID"
        .ColKey(22) = "EName"
        .ColKey(23) = "SMSYN"
        .ColKey(24) = "OldNNewClss"
        .ColKey(25) = "GunMoolMngNo"
        .ColKey(26) = "Address1"
        .ColKey(27) = "Address2"
        .ColKey(28) = "AddressAssist"
        '-----------------------------------------------------------------------


        .Redraw = flexRDDirect
    End With

    With grdMenu
        .Cols = 9
        Call SetVSFlexGrid(grdMenu)

        .Rows = 1
        .FixedCols = 0
        .FixedRows = 1
        .ExtendLastCol = False

        .OutlineBar = flexOutlineBarSimpleLeaf
        .OutlineCol = 1
        .GridLines = flexGridNone
        .Editable = flexEDNone
        
        .TextArray(0) = "":             .ColWidth(0) = 0:               .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "�޴���":       .ColWidth(1) = 3930:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "��뱸��":     .ColWidth(2) = 900:             .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "���":         .ColWidth(3) = 600:             .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "�߰�":         .ColWidth(4) = 600:             .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "����":         .ColWidth(5) = 600:             .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "����":         .ColWidth(6) = 600:             .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "ParentID":     .ColWidth(7) = 0:               .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "MenuID":       .ColWidth(8) = 0:               .ColAlignment(8) = flexAlignCenterCenter
        
        .ColDataType(2) = flexDTBoolean
        .ColDataType(3) = flexDTBoolean
        .ColDataType(4) = flexDTBoolean
        .ColDataType(5) = flexDTBoolean
        .ColDataType(6) = flexDTBoolean
    End With
    
    With grdProcess
        .Cols = 8
        .Redraw = flexRDNone
        Call SetVSFlexGrid(grdProcess)
        .ScrollBars = flexScrollBarVertical
        
        .Rows = 1
        
        .TextArray(0) = "":             .ColWidth(0) = 0:               .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "������":       .ColWidth(1) = 1800:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "������":       .ColWidth(2) = 300:             .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "�����":       .ColWidth(3) = 1200:            .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "ȣ��":         .ColWidth(4) = 1100:             .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "ȣ��":         .ColWidth(5) = 300:             .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "�����ڵ�":     .ColWidth(6) = 0:               .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "MachineID":    .ColWidth(7) = 0:               .ColAlignment(7) = flexAlignCenterCenter
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
    
        .Redraw = flexRDDirect
    End With
    
    With grdMachine
        .Cols = 4
        .Redraw = flexRDNone
        Call SetVSFlexGrid(grdMachine)

        .Rows = 1
        
        .TextArray(0) = "":             .ColWidth(0) = 0:               .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "�����":       .ColWidth(1) = 2000:            .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "ȣ��":         .ColWidth(2) = 1100:            .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "ȣ��":         .ColWidth(3) = 0:             .ColAlignment(3) = flexAlignCenterCenter
       
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub FillGrid(Optional sDepart As String = "")
    Dim oPerson As PlusLib2.CPerson
    Dim rs      As Recordset
    
    Dim i%, lCurRow%
    Dim lsAdditemStr As String
    On Error GoTo ErrHandler

    Set oPerson = New PlusLib2.CPerson
    oPerson.Connection = g_adoCon

    Set rs = oPerson.GetPerson(sDepart)
    Set oPerson = Nothing

    m_bloading = True

    With grdData
        .Redraw = flexRDNone

        lCurRow = IIf(.Row > .FixedRows - 1, .Row, .FixedRows)
        .Rows = .FixedRows

        i = 1
        Do Until rs.EOF
''            'S_201312_��������_99 �� ���� ����-OLD�ҽ�
''            .AddItem CStr(i) & vbTab & IIf(Len(CheckNull(rs!Enddate)) > 0, "��", "") & vbTab & rs!Name & vbTab & CStr(rs!PersonID) & vbTab & _
''                rs!UserID & vbTab & CheckNull(rs!Depart) & vbTab & CheckNull(rs!Duty) & vbTab & MakeDate(DF_LONG, CheckNull(rs!StartDate)) & vbTab & _
''                Format(CheckNull(rs!RegistID), "######-#######") & vbTab & CheckNull(rs!Phone) & vbTab & _
''                IIf(CheckNull(rs!SolarClss) = "0", "��", "��") & "," & MakeDate(DF_LONG, CheckNull(rs!BirthDay)) & vbTab & _
''                CheckNull(rs!Address1) & vbTab & CheckNull(rs!Address2) & vbTab & _
''                CheckNull(rs!ZipCode) & vbTab & CheckNull(rs!Remark) & vbTab & MakeDate(DF_LONG, CheckNull(rs!Enddate)) & vbTab & _
''                CheckNull(rs!Password) & vbTab & CheckNull(rs!DepartID) & vbTab & CheckNull(rs!DutyID) & vbTab & CheckNull(rs!HandPhone) & _
''                vbTab & CheckNull(rs!Email) & vbTab & CheckNull(rs!TeamID)

                 'S_201312_��������_99 �� ���� ����-NEW�ҽ�
                lsAdditemStr = CStr(i)                                                                                          '0)Row ��
                lsAdditemStr = lsAdditemStr & vbTab & IIf(Len(CheckNull(rs!EndDate)) > 0, "��", "")                             '1)��翩��
                lsAdditemStr = lsAdditemStr & vbTab & rs!Name                                                                   '2)����
                lsAdditemStr = lsAdditemStr & vbTab & CStr(rs!PersonID)                                                         '3)���ID
                lsAdditemStr = lsAdditemStr & vbTab & rs!UserID                                                                 '4)����ڹ�ȣ
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Depart)                                                      '5)�μ�
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Duty)                                                        '6)��å
                lsAdditemStr = lsAdditemStr & vbTab & MakeDate(DF_LONG, CheckNull(rs!StartDate))                                '7)�Ի�����
                lsAdditemStr = lsAdditemStr & vbTab & Format(CheckNull(rs!RegistID), "######-#######")                          '8)�ֹε�Ϲ�ȣ
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Phone)                                                       '9)��ȭ��ȣ
                lsAdditemStr = lsAdditemStr & vbTab & IIf(CheckNull(rs!SolarClss) = "0", "��", "��") & "," & MakeDate(DF_LONG, CheckNull(rs!BirthDay))  '10)�������
                
                'S_201312_��������_99 �� ���� ����(OLD:rs!Address1)
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!AddressJiBun1)                                               '11)�����ּ�1
                'S_201312_��������_99 �� ���� ����(OLD:rs!Address2)
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!AddressJiBun2)                                               '12)�����ּ�2
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!ZipCode)                                                     '13)�����ȣ
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Remark)                                                      '14)������
                lsAdditemStr = lsAdditemStr & vbTab & MakeDate(DF_LONG, CheckNull(rs!EndDate))                                  '15)�������
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Password)                                                    '16)Password
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!DepartID)                                                    '17)DepartID
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!DutyID)                                                      '18)DutyID
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!HandPhone)                                                   '19)�ڵ���
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Email)                                                       '20)EMail
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!TeamID)                                                      '21)�۾���
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!EName)                                                       '22)�����̸�
                lsAdditemStr = lsAdditemStr & vbTab & IIf(CheckNull(rs!SMSYN) = "Y", "Y", "N")                                  '23)�������۴�󿩺�
                'S_201312_��������_99 �� ���� �߰�-----------------------------------------
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!OldNNewClss)                                                 '24)�ּұ���
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!GunMoolMngNo)                                                '25)�ǹ�������ȣ
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Address1)                                                    '26)���θ��ּ�1
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Address2)                                                    '27)���θ��ּ�1
                lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!AddressAssist)                                               '28)���θ� ���� �ּ�
                '---------------------------------------------------------------------
                        
                .AddItem lsAdditemStr
                
                

            i = i + 1
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        m_bloading = False

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways

            .Row = IIf(.Rows > lCurRow, lCurRow, .Rows - 1)
            .Col = .FixedCols
            .ColSel = .Cols - 1

            Call ShowData
        Else
            .HighLight = flexHighlightNever

            Call ClearData
        End If

        If Len(Trim(txtSearch)) > 0 Then
            Call txtSearch_Change
        Else
            Call ChangeScroll
        End If

        .Redraw = flexRDDirect
    End With

    lblCount.Caption = LoadResString(250) & grdData.Rows - 1 & "  ��"

    Exit Sub

ErrHandler:
    m_bloading = False
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Sub

Private Sub ShowData()
    Dim oPerson As PlusLib2.CPerson
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler

    If m_bloading Then Exit Sub

    With grdData
''        'S_201312_��������_99 �� ���� ����-OLD�ҽ�
''        txtCode = .TextMatrix(.Row, 3)
''        txtName = .TextMatrix(.Row, 2)
''        txtUserID = .TextMatrix(.Row, 4)
''        cboDepart.ListIndex = FindComboBox(cboDepart, CLng(.TextMatrix(.Row, 17)))
''        cboDuty.ListIndex = FindComboBox(cboDuty, CLng(.TextMatrix(.Row, 18)))
''        mskStartDate = MakeDate(DF_SHORT, .TextMatrix(.Row, 7))
''        mskRegistID = .TextMatrix(.Row, 8)
''        txtTelePhone = .TextMatrix(.Row, 9)
''        cboSolarClss.ListIndex = IIf(Left(.TextMatrix(.Row, 10), 1) = "��", 0, 1)
''        mskBirthday = MakeDate(DF_SHORT, Mid(.TextMatrix(.Row, 10), 3, 13))
''        txtAddress(0) = .TextMatrix(.Row, 11)
''        txtAddress(1) = .TextMatrix(.Row, 12)
''        mskZipCode = .TextMatrix(.Row, 13)
''        txtRemark = .TextMatrix(.Row, 14)
''        chkEnd.Value = IIf(Len(.TextMatrix(.Row, 1)) > 0, vbChecked, vbUnchecked)
''        mskEndDate = .TextMatrix(.Row, 15)
''        txtPassWord = .TextMatrix(.Row, 16)
''        txtHandPhone = .TextMatrix(.Row, 19)
''        txtEMail = .TextMatrix(.Row, 20)
''        cboTeam.ListIndex = IIf(Len(Trim(.TextMatrix(.Row, 21))) = 0, 0, .TextMatrix(.Row, 21) - 1)

        'S_201312_��������_99 �� ���� ����-NEW�ҽ�
        txtCode = .TextMatrix(.Row, .ColIndex("PersonID"))                                                 '�����ȣ[3]
        txtName = .TextMatrix(.Row, .ColIndex("Name"))                                                     '�̸�[2]
        
        txtename = .TextMatrix(.Row, .ColIndex("EName"))                                                                   '�����̸�
        
        
        txtUserID = .TextMatrix(.Row, .ColIndex("UserID"))                                                 '�����ID[4]
        cboDepart.ListIndex = FindComboBox(cboDepart, CLng(.TextMatrix(.Row, .ColIndex("DepartID"))))       '�μ�[17]
        cboDuty.ListIndex = FindComboBox(cboDuty, CLng(.TextMatrix(.Row, .ColIndex("DutyID"))))             '��å[18]
        mskStartDate = MakeDate(DF_SHORT, .TextMatrix(.Row, .ColIndex("StartDate")))                        '�Ի�����[7]
        mskRegistID = .TextMatrix(.Row, .ColIndex("RegistID"))                                              '�ֹε�Ϲ�ȣ[8]
        txtTelePhone = .TextMatrix(.Row, .ColIndex("Phone"))                                                '��ȭ��ȣ[9]
        cboSolarClss.ListIndex = IIf(Left(.TextMatrix(.Row, .ColIndex("BirthDay")), 1) = "��", 0, 1)        '���/���±���[10]
        mskBirthday = MakeDate(DF_SHORT, Mid(.TextMatrix(.Row, .ColIndex("BirthDay")), 3, 13))              '�������[10]
        mskZipCode = .TextMatrix(.Row, .ColIndex("ZipCode"))                                                '�����ȣ[13]
        'S_201312_��������_99 �� ���� �߰�-----------------------------------------------------------------
        If .TextMatrix(.Row, .ColIndex("OldNNewClss")) = "0" Then
            optOldNNew(0).Value = True                                                                      '���θ��ּҼ���[24]
        Else
            optOldNNew(1).Value = True                                                                      '�����ּ�
        End If
        
        txtGunMoolMngNo.Text = .TextMatrix(.Row, .ColIndex("GunMoolMngNo"))                                 '�ǹ����� �����ĺ���ȣ[25]
        txtAddress1.Text = .TextMatrix(.Row, .ColIndex("Address1"))                                         ' �ּ�-���θ�[26]
        txtAddress2.Text = .TextMatrix(.Row, .ColIndex("Address2"))                                         '�ּ�2-���θ�[27]
        txtAddressAssist.Text = .TextMatrix(.Row, .ColIndex("AddressAssist"))                               '���θ� �����ּ�[28]
        '------------------------------------------------------------------------------------------------
        
        txtAddress(0) = .TextMatrix(.Row, .ColIndex("AddressJiBun1"))                                       '�����ּ�1[11]
        txtAddress(1) = .TextMatrix(.Row, .ColIndex("AddressJiBun2"))                                       '�����ּ�2[12]

        txtRemark = .TextMatrix(.Row, .ColIndex("Remark"))                                                  '���[14]
        chkEnd.Value = IIf(Len(.TextMatrix(.Row, .ColIndex("DspEnd"))) > 0, vbChecked, vbUnchecked)         '��翩��[1]
        mskEndDate = .TextMatrix(.Row, .ColIndex("EndDate"))                                                '�������[15]
        txtPassWord = .TextMatrix(.Row, .ColIndex("Password"))                                              '��ȣ[16]
        txtHandPhone = .TextMatrix(.Row, .ColIndex("HandPhone"))                                            '�ڵ���[19]
        txtEMail = .TextMatrix(.Row, .ColIndex("Email"))                                                    '�̸���[20]
        'cboTeam.ListIndex = IIf(Len(Trim(.TextMatrix(.Row, 21))) = 0, 0, CInt(.TextMatrix(.Row, 21)))
        If Len(.TextMatrix(.Row, .ColIndex("TeamID"))) = 0 Then                                             '�۾���[21]
            cboTeam.ListIndex = 0
        Else
            cboTeam.ListIndex = IIf(Len(Trim(.TextMatrix(.Row, .ColIndex("TeamID")))) = 0, 0, .TextMatrix(.Row, .ColIndex("TeamID")) - 1)
        End If

        chkSMSYN.Value = IIf(.TextMatrix(.Row, .ColIndex("SMSYN")) = "Y", vbChecked, vbUnchecked)         '�������۴�󿩺�

    End With
    
    Set oPerson = New PlusLib2.CPerson
    oPerson.Connection = g_adoCon
    
    Set rs = oPerson.GetPersonMachine(txtCode)
    Set oPerson = Nothing
   
   
    With grdProcess
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs!Process & vbTab & " " & vbTab & rs!Machine & vbTab & rs!MachineNO & vbTab & _
                        " " & vbTab & rs!ProcessID & vbTab & rs!machineid
            
            .Cell(flexcpPicture, .Rows - 1, 2) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, .Rows - 1, 2) = flexPicAlignCenterCenter
            
            .Cell(flexcpPicture, .Rows - 1, 5) = LoadResPicture("B_FIND", vbResBitmap)
            .Cell(flexcpPictureAlignment, .Rows - 1, 5) = flexPicAlignCenterCenter
            
            rs.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    
    End With
    
    rs.Close
    Set rs = Nothing

    Call MakeMenu(txtUserID)
Exit Sub

ErrHandler:
    Set oPerson = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
    
End Sub

Private Function CheckData() As Boolean
    Dim i%

    CheckData = True

    If Len(Trim(txtName)) = 0 Then
        MsgBox "'����'�� �Է��ϼž� �մϴ�.", vbInformation
        txtName.SetFocus
        GoTo FailedCheckData
    End If
    
    If Len(Trim(txtename)) = 0 Then
        MsgBox "'�����̸�'�� �Է��ϼž� �մϴ�.", vbInformation
        txtName.SetFocus
        GoTo FailedCheckData
    End If

    If cboDepart.ListIndex < 0 Then
        MsgBox "'�μ�'�� �����ϼž� �մϴ�.", vbInformation
        cboDepart.SetFocus
        GoTo FailedCheckData
    End If

    If Len(Trim(txtPassWord)) = 0 Then
        MsgBox "'�н�����'�� �Է��ϼž� �մϴ�.", vbInformation
        txtUserID.SetFocus
        GoTo FailedCheckData
    End If

    If chkEnd = vbChecked And Len(Trim(mskEndDate.Text)) <> 8 Then
        MsgBox "������ڸ� �߸� �Է� �ϼ̽��ϴ�", vbInformation, "������"
        mskEndDate.SetFocus
        GoTo FailedCheckData
    End If
    
    With grdData
        For i = 1 To .Rows - 1
            If txtUserID = .TextMatrix(i, 4) And m_sFlag = ID_ADDNEW Then
                MsgBox "������ ID�� �ֽ��ϴ�. �ٽ� �Է��Ͻʽÿ�.", vbInformation
                txtUserID.SetFocus
                GoTo FailedCheckData
            End If
        Next i
    End With

    With grdData
        If Len(txtCode) > 0 Then
            For i = 1 To .Rows - 1
                If txtCode = .TextMatrix(i, 3) And m_sFlag = ID_ADDNEW Then
                    MsgBox "������ ���ID�� �ֽ��ϴ�. �ٽ� �Է��Ͻʽÿ�.", vbInformation
                    txtCode.SetFocus
                    GoTo FailedCheckData
                End If
            Next i
        End If
    End With

    Exit Function

FailedCheckData:
    CheckData = False
    Exit Function
End Function

Private Function SaveData() As Boolean
    Dim oPerson  As PlusLib2.CPerson
    Dim stData   As PlusLib2.TPerson
    Dim oMenu    As PlusLib2.CMenu
    Dim stMenu() As PlusLib2.TUSERMENU
    Dim stMachine() As PlusLib2.TPersonMachine
    Dim iLoop%, nCnt%, i%, nCount%
    Dim bCheck As Boolean
    
    On Error GoTo ErrHandler

    MousePointer = vbHourglass
    
    nCnt = 0
    ReDim Preserve stMenu(nCnt)
    With grdMenu
        For iLoop = .FixedRows + 1 To .Rows - 1
        
            ''' ���� ����~~
            If val(Left(Right(.TextMatrix(iLoop, 1), 5), 4)) Mod 100 = 0 Then
                bCheck = IIf(.Cell(flexcpPicture, iLoop, 2) = ImgCheck, True, False)
            Else
                bCheck = IIf(.Cell(flexcpChecked, iLoop, 2) = 1, True, False)
            End If


            If bCheck Then
                stMenu(nCnt).sPersonID = IIf(m_sFlag = ID_UPDATE, txtCode, "")
                stMenu(nCnt).sMenuID = Left(Right(.TextMatrix(iLoop, 1), 5), 4)
                stMenu(nCnt).nSeq = nCnt
                stMenu(nCnt).sParentID = IIf(.TextMatrix(iLoop, 7) = "", 0, .TextMatrix(iLoop, 7))
                stMenu(nCnt).nLevel = .RowOutlineLevel(iLoop) - 1
                
                If val(stMenu(nCnt).sMenuID) Mod 100 = 0 Then
                    stMenu(nCnt).sPrintClss = IIf(.Cell(flexcpPicture, iLoop, 3) = ImgCheck, "*", "")
                    'stMenu(nCnt).sSelectClss = IIf(.Cell(flexcpPicture, iLoop, 3) = ImgCheck, "*", "")
                    stMenu(nCnt).sAddNewClss = IIf(.Cell(flexcpPicture, iLoop, 3) = ImgCheck, "*", "")
                    stMenu(nCnt).sUpdateClss = IIf(.Cell(flexcpPicture, iLoop, 3) = ImgCheck, "*", "")
                    stMenu(nCnt).sDeleteClss = IIf(.Cell(flexcpPicture, iLoop, 3) = ImgCheck, "*", "")
                Else
                    
                    stMenu(nCnt).sPrintClss = IIf(.Cell(flexcpChecked, iLoop, 3) = 1, "*", "")
                    'stMenu(nCnt).sSelectClss = IIf(.Cell(flexcpChecked, iLoop, 3) = 1, "*", "")
                    stMenu(nCnt).sAddNewClss = IIf(.Cell(flexcpChecked, iLoop, 4) = 1, "*", "")
                    stMenu(nCnt).sUpdateClss = IIf(.Cell(flexcpChecked, iLoop, 5) = 1, "*", "")
                    stMenu(nCnt).sDeleteClss = IIf(.Cell(flexcpChecked, iLoop, 6) = 1, "*", "")
                    
                End If
                
                nCnt = nCnt + 1
                ReDim Preserve stMenu(nCnt)
            End If
            
        Next iLoop
    End With
    

    With stData
        .sPersonID = Trim(txtCode)                      '���ID
        .sUserID = txtUserID                            '�����ID
        .sName = Trim(txtName)                          '�ѱۼ���
        .sEname = Trim(txtename)                        '�����̸�
        .sDepartID = Format(CStr(cboDepart.ItemData(cboDepart.ListIndex)), FORMAT_DEPARTID)     '�μ�
        .sDutyID = Format((cboDuty.ItemData(cboDuty.ListIndex)), FORMAT_DUTYID)                 '��å
        .sStartDate = mskStartDate                      '�Ի�����
        .sPassword = Trim(txtPassWord)                  '��ȣ
        .sRegistID = mskRegistID                        '�ֹι�ȣ
        .sSolarClss = cboSolarClss.ListIndex            '���� ����
        .sBirthDay = mskBirthday                        '�������
        .sHandPhone = txtHandPhone                      '�޴���
        .sPhone = txtTelePhone                          '��ȭ��ȣ
        .sZipCode = IIf(Len(mskZipCode) = 0, "", Left(mskZipCode, 3) + "-" + Right(mskZipCode, 3))      '�����ȣ
        'S_201312_��������_99 �� ���� �߰�-------------------------------------------------------
        .sOldNNewClss = IIf(optOldNNew(0).Value = True, "0", "1")    '���θ�,�����ּ� ���� 0:���θ�, 1:����
        .sGunMoolMngNo = IIf(optOldNNew(0).Value = True, txtGunMoolMngNo.Text, "")        '�ǹ����� �����ĺ���ȣ
        .sAddress1 = txtAddress1.Text        ' ���θ� �ּ�1
        .sAddress2 = txtAddress2.Text        '���θ� �ּ�2
        .sAddressAssist = txtAddressAssist.Text         '���θ� ���� �ּ�
        '----------------------------------------------------------------------------------------
        'S_201312_��������_99 �� ���� ����(OLD:.sAddress1)
        .sAddressJiBun1 = Trim(txtAddress(0))
        'S_201312_��������_99 �� ���� ����(OLD:.sAddress2)
        .sAddressJiBun2 = Trim(txtAddress(1))
        .sRemark = Trim(txtRemark)                  '���
        .SendDate = mskEndDate                      '�������
        .sEMail = txtEMail                          '�̸���
        .sTeamID = Format(cboTeam.ListIndex + 1, "00")      '�۾���
        .sSMSYN = IIf(chkSMSYN.Value = vbChecked, "Y", "N")     '�������۴�󿩺�
        
    End With

    ' ��������
    nCount = 0
    ReDim Preserve stMachine(nCount)
    With grdProcess
        For iLoop = .FixedRows To .Rows - 1
            stMachine(nCount).sPersonID = stData.sPersonID
            stMachine(nCount).sProcessID = .TextMatrix(iLoop, 6)
            stMachine(nCount).sMachineID = Format(.TextMatrix(iLoop, 7), "00")
                            
            nCount = nCount + 1
            ReDim Preserve stMachine(nCount)
                    
        Next iLoop
    End With
    
    Set oPerson = New PlusLib2.CPerson
    oPerson.Connection = g_adoCon
    oPerson.UserName = g_sUserName

    If m_sFlag = ID_ADDNEW Then
        SaveData = oPerson.AddNewPerson(stData, stMenu, nCnt, stMachine, nCount)
    ElseIf m_sFlag = ID_UPDATE Then
        SaveData = oPerson.UpdatePerson(stData, stMenu, nCnt, stMachine, nCount)
    End If

    Set oPerson = Nothing
        
    MousePointer = vbDefault

    Exit Function

ErrHandler:
    MousePointer = vbDefault
    SaveData = False

    Call ErrorBox(Err.Number, Err.Source, Err.Description)
    Err.Clear
End Function

Private Function DeleteData() As Boolean
    Dim oPerson As PlusLib2.CPerson
    Dim oMenu   As PlusLib2.CMenu

    On Error GoTo ErrHandler
    
    MousePointer = vbHourglass

    Set oPerson = New PlusLib2.CPerson
    oPerson.Connection = g_adoCon
    DeleteData = oPerson.DeletePerson(Format((txtCode.Text), FORMAT_PERSONID), MakeDate(DF_SHORT, Now))
    
    Set oPerson = Nothing

    If Not DeleteData Then Exit Function


    MousePointer = vbDefault
    Exit Function
ErrHandler:
    MousePointer = vbDefault
    DeleteData = False
    
    Call ErrorBox(Err.Number, "Person.DeleteData", Err.Description)
End Function


Private Sub ClearData()
    txtCode = ""
    txtName = ""
    txtename = ""               '�����̸�
    cboDepart.ListIndex = 0
    cboDuty.ListIndex = 0
    mskStartDate = ""
    mskEndDate = ""
    txtUserID = ""
    txtPassWord = ""
    mskRegistID = ""
    cboSolarClss.ListIndex = 0
    mskBirthday = ""
    txtHandPhone = ""
    txtTelePhone = ""
    txtAddress(0) = ""
    txtAddress(1) = ""
    mskZipCode = ""
    'S_201312_��������_99 �� ���� �߰�---------------------------------------
    optOldNNew(0).Value = True     '���θ��ּҼ���
    txtGunMoolMngNo.Text = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtAddressAssist.Text = ""
    '--------------------------------------------------------------------
    
    txtRemark = ""
    
    chkEnd.Value = vbUnchecked
    
    chkSMSYN.Value = vbUnchecked        '�������� ��󿩺�
End Sub

Private Sub MakeMenu(Optional sUserID As String)
    Dim oMenu As PlusLib2.CMenu
    Dim rs    As ADODB.Recordset
    Dim sMenu$
    Dim i%, irow

    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass

    Set oMenu = New PlusLib2.CMenu
    oMenu.Connection = g_adoCon

    With grdMenu
        .Redraw = flexRDNone

        .Rows = .FixedRows

        Call AddItemGridMenu("�޴����", "0000", 0, True)
        .RowOutlineLevel(.FixedRows) = 0
        
        ' ���� ��� �޴��� ��� üũ �ڽ���� �������� �̹��� �÷��� ��.
        For i = 2 To 6
            .Cell(flexcpPictureAlignment, .Rows - 1, i) = flexPicAlignCenterCenter
            .Cell(flexcpPicture, .Rows - 1, i) = imgUnCheck
        Next i

        Set rs = oMenu.GetMainMenu()

        Do While Not rs.EOF
            sMenu = rs!Menu & " (" & rs!MenuID & ")"
            If CInt(rs!MenuID) Mod 100 = 0 Then
                Call AddItemGridMenu(sMenu, rs!MenuID, CheckNull(rs!ParentID), True)

                If IsNull(rs!ParentID) Or rs!ParentID = 0 Then
                    .RowOutlineLevel(.Rows - 1) = 1
                    
                ElseIf rs!ParentID < 1000 Then
                    .RowOutlineLevel(.Rows - 1) = 2
                Else
                    .RowOutlineLevel(.Rows - 1) = 3
                End If
                
                For i = 2 To 6
                    .Cell(flexcpPictureAlignment, .Rows - 1, i) = flexPicAlignCenterCenter
                    .Cell(flexcpPicture, .Rows - 1, i) = imgUnCheck
                Next i
            Else
                Call AddItemGridMenu(sMenu, rs!MenuID, CheckNull(rs!ParentID))
                If Not (IsNull(rs!ParentID) Or rs!ParentID = 0) Then
                    .RowOutlineLevel(.Rows - 1) = 4
                Else
                    .RowOutlineLevel(.Rows - 1) = 1
                End If
                
            End If

            rs.MoveNext
            
        Loop
        rs.Close

        If Len(sUserID) > 0 Then
            Set rs = oMenu.GetUserMenu(sUserID)

            Do While Not rs.EOF
                Call SetItemGridMenu(rs!MenuID, True, IIf(rs!PrintClss = "*", True, False), IIf(rs!AddNewClss = "*", True, False), IIf(rs!UpdateClss = "*", True, False), _
                    IIf(rs!DeleteClss = "*", True, False))
                   
                rs.MoveNext
            Loop

            rs.Close
        End If
        Set oMenu = Nothing

        .Row = .FixedRows

        .Redraw = flexRDDirect
        Call ChangeMenuScroll
    End With
    Set rs = Nothing
    Set oMenu = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oMenu = Nothing
    Screen.MousePointer = vbDefault

    Call ErrorBox(Err.Number, "frmPerson.MakeMenu", Err.Description)
End Sub

Private Sub AddItemGridMenu(sMenu As String, sMenuID As String, sParentID As String, Optional bSubTotal As Boolean = False, Optional bVisible As Boolean = False, _
    Optional bAddNew As Boolean = False, Optional bUpdate As Boolean = False, Optional bDelete As Boolean = False, Optional bOutput As Boolean = False)
    Dim irow%

    With grdMenu
        .AddItem "" & vbTab & sMenu & vbTab & bVisible & vbTab & bAddNew & vbTab & bUpdate & vbTab & bDelete & vbTab & bOutput & vbTab & sParentID

        irow = .Rows - 1

        .RowData(irow) = sMenuID
        If bSubTotal Then
            .IsSubtotal(irow) = True
            .Cell(flexcpPicture, irow, 1) = imgFolder
        Else
            .Cell(flexcpPicture, irow, 1) = imgItem
        End If
    End With
End Sub

Private Sub SetItemGridMenu(sMenuID As String, bVisible As Boolean, bAddNew As Boolean, bUpdate As Boolean, bDelete As Boolean, bOutput As Boolean)
    Dim irow%

    With grdMenu
        irow = .FindRow(sMenuID)
        If irow < 0 Then Exit Sub
        If Not CInt(sMenuID) Mod 100 = 0 Then
        
            .Cell(flexcpChecked, irow, 2) = IIf(bVisible, True, False) '��뱸��
            .Cell(flexcpChecked, irow, 3) = IIf(bAddNew, True, False)  '�߰�
            .Cell(flexcpChecked, irow, 4) = IIf(bUpdate, True, False)   '����
            .Cell(flexcpChecked, irow, 5) = IIf(bDelete, True, False)   '����
            .Cell(flexcpChecked, irow, 6) = IIf(bOutput, True, False)   '����
        Else
            ' �����׷� �޴��� ��� �̹����� �������� �Է�
            .Cell(flexcpPicture, irow, 2) = IIf(bVisible, ImgCheck, imgUnCheck)
            .Cell(flexcpPicture, irow, 3) = IIf(bAddNew, ImgCheck, imgUnCheck)
            .Cell(flexcpPicture, irow, 4) = IIf(bUpdate, ImgCheck, imgUnCheck)
            .Cell(flexcpPicture, irow, 5) = IIf(bDelete, ImgCheck, imgUnCheck)
            .Cell(flexcpPicture, irow, 6) = IIf(bOutput, ImgCheck, imgUnCheck)
        End If
            
    End With
End Sub

Private Sub ChangeScroll()
    Dim lRows As Long

    lRows = GetVisibleVSGridRowCount(grdData)

    With grdData
        If lRows > LIMIT_ROW1 Then
            .ColWidth(4) = LIMIT_WIDTH1 - 240
        Else
            .ColWidth(4) = LIMIT_WIDTH1
        End If
    End With

    If lRows = 0 Then
        Call ClearData
        cmdOperate(ID_UPDATE).Enabled = False
        cmdOperate(ID_DELETE).Enabled = False
        cmdPrint.Enabled = False
    Else
        Call ShowData
        cmdOperate(ID_UPDATE).Enabled = True
        cmdOperate(ID_DELETE).Enabled = True
        cmdPrint.Enabled = True
    End If
End Sub

Private Sub ChangeMenuScroll()
    With grdMenu
        .ColWidth(1) = LIMIT_WIDTH2 - IIf(GetVisibleVSGridRowCount(grdMenu) > LIMIT_ROW2, 240, 0)
    End With
End Sub


