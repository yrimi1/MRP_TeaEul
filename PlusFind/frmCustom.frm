VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCustom 
   BackColor       =   &H8000000A&
   Caption         =   "�ŷ�ó ����"
   ClientHeight    =   8310
   ClientLeft      =   3060
   ClientTop       =   2895
   ClientWidth     =   11865
   Icon            =   "frmCustom.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11865
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   6495
      Left            =   15
      TabIndex        =   59
      Top             =   1005
      Width           =   3495
      _cx             =   6165
      _cy             =   11456
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
   Begin VB.CommandButton cmdOperate 
      Cancel          =   -1  'True
      Caption         =   "���(&C)"
      Height          =   810
      Index           =   4
      Left            =   8550
      MousePointer    =   99  '����� ����
      Style           =   1  '�׷���
      TabIndex        =   43
      ToolTipText     =   "�ڷ� ���"
      Top             =   135
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdOperate 
      Caption         =   "����(&S)"
      Height          =   810
      Index           =   3
      Left            =   7695
      MousePointer    =   99  '����� ����
      Style           =   1  '�׷���
      TabIndex        =   42
      ToolTipText     =   "�ڷ� ����"
      Top             =   135
      Visible         =   0   'False
      Width           =   840
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10125
      TabIndex        =   55
      Top             =   7590
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      ���(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   915
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1614
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optSize 
         Caption         =   "���"
         Height          =   330
         Index           =   0
         Left            =   2745
         Style           =   1  '�׷���
         TabIndex        =   5
         Top             =   480
         Width           =   645
      End
      Begin VB.OptionButton optSize 
         Caption         =   "��"
         Height          =   330
         Index           =   1
         Left            =   2730
         Style           =   1  '�׷���
         TabIndex        =   4
         Top             =   90
         Value           =   -1  'True
         Width           =   645
      End
      Begin Threed.SSCommand cmdAll 
         Height          =   330
         Left            =   1920
         TabIndex        =   3
         Top             =   465
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
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   105
         TabIndex        =   2
         Top             =   480
         Width           =   1755
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   25
         Left            =   105
         TabIndex        =   1
         Top             =   135
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "��ȣ �˻���"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlMain 
      Height          =   7470
      Left            =   3555
      TabIndex        =   6
      Top             =   45
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   13176
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&D)"
         Height          =   810
         Index           =   2
         Left            =   7440
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   9
         ToolTipText     =   "�ڷ� ����"
         Top             =   90
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "�߰�(&A)"
         Height          =   810
         Index           =   0
         Left            =   5850
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   7
         ToolTipText     =   "�ڷ� �߰�"
         Top             =   90
         Width           =   780
      End
      Begin Threed.SSPanel pnlEdit 
         Height          =   6435
         Left            =   75
         TabIndex        =   44
         Top             =   975
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   11351
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Frame fraAddress 
            Caption         =   "�ּ�"
            Height          =   1905
            Left            =   60
            TabIndex        =   76
            Top             =   3090
            Width           =   8025
            Begin VB.Frame fraJiBun 
               Caption         =   "����"
               Height          =   825
               Left            =   1950
               TabIndex        =   79
               Top             =   1020
               Width           =   6015
               Begin PlusFind2.WizText txtAddressJiBun1 
                  Height          =   300
                  Left            =   60
                  TabIndex        =   33
                  Top             =   180
                  Width           =   5865
                  _ExtentX        =   10345
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
               Begin PlusFind2.WizText txtAddressJiBun2 
                  Height          =   300
                  Left            =   60
                  TabIndex        =   34
                  Top             =   480
                  Width           =   5865
                  _ExtentX        =   10345
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
            Begin VB.Frame fraDoro 
               Caption         =   "���θ�"
               Height          =   855
               Left            =   1950
               TabIndex        =   78
               Top             =   150
               Width           =   6015
               Begin VB.TextBox txtGunMoolMngNo 
                  Enabled         =   0   'False
                  Height          =   270
                  Left            =   4980
                  TabIndex        =   80
                  TabStop         =   0   'False
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin PlusFind2.WizText txtAddress1 
                  Height          =   300
                  Left            =   60
                  TabIndex        =   30
                  Top             =   180
                  Width           =   5865
                  _ExtentX        =   10345
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
               Begin PlusFind2.WizText txtAddress2 
                  Height          =   300
                  Left            =   60
                  TabIndex        =   31
                  Top             =   480
                  Width           =   3765
                  _ExtentX        =   6641
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
               Begin PlusFind2.WizText txtAddressAssist 
                  Height          =   300
                  Left            =   3840
                  TabIndex        =   32
                  Top             =   480
                  Width           =   2085
                  _ExtentX        =   3678
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
               TabIndex        =   77
               Top             =   150
               Width           =   1875
               Begin VB.OptionButton optOldNNew 
                  Caption         =   "����"
                  Height          =   225
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   27
                  Top             =   120
                  Width           =   675
               End
               Begin VB.OptionButton optOldNNew 
                  Caption         =   "���θ�"
                  Height          =   225
                  Index           =   0
                  Left            =   60
                  TabIndex        =   26
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   855
               End
            End
            Begin Threed.SSCommand cmdFind 
               Height          =   315
               Left            =   1080
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   570
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               _Version        =   196609
               ButtonStyle     =   3
               Outline         =   0   'False
            End
            Begin MSMask.MaskEdBox mskZipCode 
               Height          =   300
               Left            =   60
               TabIndex        =   28
               Top             =   570
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               Mask            =   "###-###"
               PromptChar      =   "_"
            End
         End
         Begin VB.Frame fraID 
            Caption         =   " ���ͳ� �α��� ���� "
            Height          =   945
            Left            =   5220
            TabIndex        =   57
            Top             =   2160
            Width           =   2865
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   8
               Left            =   90
               TabIndex        =   50
               Top             =   255
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "�� �� ��"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   9
               Left            =   90
               TabIndex        =   51
               Top             =   570
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "��й�ȣ"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin PlusFind2.WizText txtUserID 
               Height          =   300
               Left            =   1320
               TabIndex        =   24
               Top             =   240
               Width           =   1425
               _ExtentX        =   2514
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
            Begin PlusFind2.WizText txtUserPassword 
               Height          =   300
               Left            =   1320
               TabIndex        =   25
               Top             =   570
               Width           =   1425
               _ExtentX        =   2514
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
         Begin VB.ComboBox cboTrade 
            Height          =   300
            Left            =   6465
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   21
            Top             =   825
            Width           =   1485
         End
         Begin MSMask.MaskEdBox mskCustomNO 
            Height          =   300
            Left            =   6450
            TabIndex        =   20
            Top             =   450
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   45
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��   ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   46
            Top             =   420
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��   ȣ"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   3
            Left            =   90
            TabIndex        =   49
            Top             =   1755
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��   ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   90
            TabIndex        =   48
            Top             =   1425
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��   ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   6
            Left            =   5205
            TabIndex        =   52
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "�� ǥ ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   7
            Left            =   5205
            TabIndex        =   53
            Top             =   450
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "����ڹ�ȣ"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   11
            Left            =   5205
            TabIndex        =   54
            Top             =   825
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "�ŷ� ����"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   16
            Left            =   90
            TabIndex        =   47
            Top             =   1080
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��ȣ (����)"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PlusFind2.WizText txtCustomID 
            Height          =   300
            Left            =   1335
            TabIndex        =   10
            Top             =   90
            Width           =   990
            _ExtentX        =   1746
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
         Begin PlusFind2.WizText txtKCustom 
            Height          =   300
            Left            =   1335
            TabIndex        =   11
            Top             =   420
            Width           =   3375
            _ExtentX        =   5953
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
         Begin PlusFind2.WizText txtECustom 
            Height          =   300
            Left            =   1335
            TabIndex        =   13
            Top             =   1080
            Width           =   3375
            _ExtentX        =   5953
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
         Begin PlusFind2.WizText txtCondition 
            Height          =   300
            Left            =   1335
            TabIndex        =   14
            Top             =   1410
            Width           =   3375
            _ExtentX        =   5953
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
         Begin PlusFind2.WizText txtCategory 
            Height          =   300
            Left            =   1335
            TabIndex        =   15
            Top             =   1740
            Width           =   3375
            _ExtentX        =   5953
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
         Begin PlusFind2.WizText txtChief 
            Height          =   300
            Left            =   6450
            TabIndex        =   19
            Top             =   90
            Width           =   1470
            _ExtentX        =   2593
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
         Begin PlusFind2.WizText txtShortCustom 
            Height          =   300
            Left            =   1335
            TabIndex        =   12
            Top             =   750
            Width           =   3375
            _ExtentX        =   5953
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
            Index           =   18
            Left            =   90
            TabIndex        =   60
            Top             =   750
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��ȣ (��Ī)"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   915
            Left            =   5220
            TabIndex        =   61
            Top             =   1170
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   1614
            _Version        =   196609
            Caption         =   " ����� "
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   17
               Left            =   90
               TabIndex        =   62
               Top             =   240
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "�� �� ��"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   19
               Left            =   90
               TabIndex        =   63
               Top             =   570
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "��ȭ ��ȣ"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin PlusFind2.WizText txtName 
               Height          =   300
               Left            =   1320
               TabIndex        =   22
               Top             =   240
               Width           =   1440
               _ExtentX        =   2540
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
            Begin PlusFind2.WizText txtPhone 
               Height          =   300
               Left            =   1320
               TabIndex        =   23
               Top             =   555
               Width           =   1440
               _ExtentX        =   2540
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
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   12
            Left            =   90
            TabIndex        =   64
            Top             =   2085
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��ǥ ��ȭ"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   5
            Left            =   90
            TabIndex        =   65
            Top             =   2400
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��ȭ ��ȣ"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PlusFind2.WizText txtPhone1 
            Height          =   300
            Left            =   1335
            TabIndex        =   16
            Top             =   2085
            Width           =   1920
            _ExtentX        =   3387
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
         Begin PlusFind2.WizText txtPhone2 
            Height          =   300
            Left            =   1335
            TabIndex        =   17
            Top             =   2400
            Width           =   1920
            _ExtentX        =   3387
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
            Index           =   13
            Left            =   90
            TabIndex        =   66
            Top             =   2730
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "�ѽ� ��ȣ"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   14
            Left            =   60
            TabIndex        =   67
            Top             =   5025
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "Ȩ ������"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   15
            Left            =   3570
            TabIndex        =   68
            Top             =   5025
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "E-MAIL"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PlusFind2.WizText txtFaxNO 
            Height          =   300
            Left            =   1335
            TabIndex        =   18
            Top             =   2730
            Width           =   1920
            _ExtentX        =   3387
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
         Begin PlusFind2.WizText txtHomepage 
            Height          =   270
            Left            =   990
            TabIndex        =   35
            Top             =   5025
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   476
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
         Begin PlusFind2.WizText txtEMail 
            Height          =   300
            Left            =   4500
            TabIndex        =   36
            Top             =   5010
            Width           =   3570
            _ExtentX        =   6297
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
         Begin Threed.SSPanel pnlEditSub 
            Height          =   1035
            Left            =   45
            TabIndex        =   69
            Top             =   5370
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   1826
            _Version        =   196609
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cboLoss 
               Height          =   300
               Left            =   1335
               TabIndex        =   37
               Top             =   75
               Width           =   2640
            End
            Begin VB.ComboBox cboSpending 
               Height          =   300
               Left            =   1335
               TabIndex        =   38
               Top             =   405
               Width           =   2640
            End
            Begin VB.ComboBox cboWorking 
               Height          =   300
               Left            =   1335
               TabIndex        =   39
               Top             =   705
               Width           =   2640
            End
            Begin VB.ComboBox cboCalc 
               Height          =   300
               Left            =   5355
               TabIndex        =   40
               Top             =   75
               Width           =   2640
            End
            Begin VB.ComboBox cboPoint 
               Height          =   300
               Left            =   5355
               TabIndex        =   41
               Top             =   375
               Width           =   2640
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   20
               Left            =   90
               TabIndex        =   70
               Top             =   690
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "������ ����"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   21
               Left            =   90
               TabIndex        =   71
               Top             =   390
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "�ҿ䷮ ����"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   22
               Left            =   90
               TabIndex        =   72
               Top             =   75
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "����/Loss"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   23
               Left            =   4095
               TabIndex        =   73
               Top             =   375
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "�Ҽ��� ó��"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   24
               Left            =   4095
               TabIndex        =   74
               Top             =   75
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "ȯ���"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   510
         Left            =   315
         TabIndex        =   56
         Top             =   210
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
         Caption         =   "����(&U)"
         Height          =   810
         Index           =   1
         Left            =   6645
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   8
         ToolTipText     =   "�ڷ� ����"
         Top             =   90
         Width           =   780
      End
   End
   Begin Threed.SSCommand cmdSelect 
      Height          =   690
      Left            =   8370
      TabIndex        =   75
      Top             =   7590
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      ����(&Q)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
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
      Height          =   210
      Left            =   105
      TabIndex        =   58
      Top             =   7800
      Width           =   3330
   End
End
Attribute VB_Name = "frmCustom"
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
'** ��������  : 2013.12.12
'**------------------------------------------------------------------------------------------------
'
'  ��û���� ID: S_201312_���_99
'  ��û��:
'  ���泯¥ : 2013.12.12
'  �۾���   : ���¿�
'  ��û���� : �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�
'  ���泻�� : ���θ�,�� �����ּ� �ɼ� ��ư �߰�
'**************************************************************************************************
Option Explicit

' �Է�/���� ���� �÷���
Private m_sFlag As String * 1
Private m_bSkip As Boolean
Private m_iSorCol As Integer

Private Const LIMIT_ROW = 23
Private Const LIMIT_WIDTH = 1870

Dim m_bSelected     As Boolean
Dim wData()

Private Sub cmdAll_Click()
    Dim iLoop As Integer
    
    With grdData
        .Redraw = flexRDNone
        
        For iLoop = .FixedRows To .Rows - 1
            .RowHidden(iLoop) = False
        Next iLoop
        .Redraw = flexRDDirect
    End With
    
    txtSearch.Text = ""
    cmdAll.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function SaveData() As Boolean
    Dim NewCustom As PlusLib2.TCustom
    Dim oCustom As PlusLib2.CCustom
    
    On Error GoTo ErrHandler
        
    If Len(Trim(txtKCustom)) = 0 Then
        MsgBox LoadResString(253), vbExclamation
        txtKCustom.SetFocus
        Exit Function
    End If
        
    With NewCustom
        If m_sFlag = ID_ADDNEW Then
            If IsNumeric(txtCustomID) Then
                .sCustomID = 0     '[1] �ŷ�ó �ڵ�
            Else
                .sCustomID = txtCustomID    '[1] �ŷ�ó �ڵ�
            End If
        Else
            .sCustomID = txtCustomID
        End If
        
        .sCustomID = IIf(Len(txtCustomID) > 0, Format(txtCustomID, "0000"), "") '[1] �ŷ�ó �ڵ�
        .sKCustom = txtKCustom          '[2] �ѱ� ��ȣ
        .sShortCustom = txtShortCustom  '[3] �ѱ� ��ȣ ����
        .sECustom = txtECustom          '[4] ������ȣ
        .sCustomNo = mskCustomNO        '[5] ����� ��ȣ
        .sChief = txtChief              '[6] ��ǥ�� ����
        .sCondition = txtCondition      '[7] ����
        .sCategory = txtCategory        '[8] ����
        .sZipCode = mskZipCode          '[9] ������ȣ
        'S_201312_���_99 �� ���� �߰�-------------------------------------------------------
        .sOldNNewClss = IIf(optOldNNew(0).Value = True, "0", "1")    '[10]  ���θ�,�����ּ� ���� 0:���θ�, 1:����
        .sGunMoolMngNo = IIf(optOldNNew(0).Value = True, txtGunMoolMngNo.Text, "")        '[11]  �ǹ����� �����ĺ���ȣ
        .sAddress1 = txtAddress1.Text        '[12] ���θ� �ּ�1
        .sAddress2 = txtAddress2.Text        '[13] ���θ� �ּ�2
        .sAddressAssist = txtAddressAssist.Text         '[14] ���θ� ���� �ּ�
        '----------------------------------------------------------------------------------------
        'S_201312_���_99 �� ���� ����(OLD:.sAddress1,txtAddress1)
        .sAddressJiBun1 = txtAddressJiBun1.Text        '[15] �ּ�1
        'S_201312_���_99 �� ���� ����(OLD:.sAddress2,txtAddress2)
        .sAddressJiBun2 = txtAddressJiBun2.Text         '[16] �ּ�2
        
        .sPhone1 = txtPhone1            '[17] ��ȭ1
        .sPhone2 = txtPhone2            '[18] ��ȭ2
        .sFaxNO = txtFaxNO              '[19] �ѽ�
        .sEMail = txtEMail              '[20] Email
        .sHomePage = txtHomepage        '[21] Ȩ ������
        .sName = txtName                '[22] ��ü ���
        .sPhone = txtPhone              '[23] ��ü��� ��ȭ
        .sTradeID = cboTrade.ItemData(cboTrade.ListIndex)     ' [24] �ŷ�����
        .sUserID = txtUserID            '[25] �ŷ�ó WebID
        .sUserPassword = txtUserPassword    '[26] �ŷ�ó WebPWD
        .sLossClss = cboLoss.ItemData(cboLoss.ListIndex)         ' [27] ����/Loss ���Կ���
        .sSpendingClss = cboSpending.ItemData(cboSpending.ListIndex) ' [28] �ҿ䷮ ������
        .sWorkingClss = cboWorking.ItemData(cboWorking.ListIndex) ' [29] ������ ������
        .sCalcClss = cboCalc.ItemData(cboCalc.ListIndex)     ' [30] Meter->Yard ȯ����
        .sPointClss = cboPoint.ItemData(cboPoint.ListIndex)   ' [31] �Ҽ��� �������
        
    End With
        
    Set oCustom = New PlusLib2.CCustom
    oCustom.Connection = adoCon
'    oCustom.UserName = g_sUserName
    If m_sFlag = ID_ADDNEW Then
        SaveData = oCustom.AddNewCustom(NewCustom)
    ElseIf m_sFlag = ID_UPDATE Then
        SaveData = oCustom.UpdateCustom(NewCustom)
    End If
    
    Set oCustom = Nothing
    Exit Function
ErrHandler:
    Set oCustom = Nothing

    Call ErrorBox(Err.Number, "Custom.SaveData", Err.Description)
End Function


'********************************************************
'* Date : 2000-12-05 (TUE)
'*
'* Description: Operate 1Button�� Index ���
'*
'********************************************************
Private Sub cmdOperate_Click(Index As Integer)
    Dim oCustom As PlusLib2.CCustom
    Dim bResult As Boolean

    On Error GoTo ErrHandler
    If optSize(0).Value Then optSize(1).Value = True

    '---------------------------------------------------------------------------
    If Index = ID_ADDNEW Then '[1] �߰�
        m_sFlag = ID_ADDNEW
        Call ChangeMode(Me, False)
        
                
        'S_201312_���_99 �� ���� �߰�-----------------------------------------------
        If optOldNNew(0).Value = True Then
            fraDoro.Enabled = True
            fraJiBun.Enabled = False
        Else
            fraDoro.Enabled = False
            fraJiBun.Enabled = True
        End If
        '-------------------------------------------------------------------------
        
        Call ClearData
        pnlMsg.Caption = LoadResString(302)
        
        txtCustomID.Locked = False
'        txtKCustom.SetFocus
    '---------------------------------------------------------------------------
    ElseIf Index = ID_UPDATE Then '[2] ����
        If grdData.Rows = grdData.FixedRows Then Exit Sub
        m_sFlag = ID_UPDATE
        Call ChangeMode(Me, False)
        
        'S_201312_���_99 �� ���� �߰�-----------------------------------------------
        If optOldNNew(0).Value = True Then
            fraDoro.Enabled = True
            fraJiBun.Enabled = False
        Else
            fraDoro.Enabled = False
            fraJiBun.Enabled = True
        End If
        '-------------------------------------------------------------------------
        
        pnlMsg.Caption = LoadResString(303)
        
        txtCustomID.Locked = True
        txtKCustom.SetFocus
    '---------------------------------------------------------------------------
    ElseIf Index = ID_DELETE Then '[3] ����
        If grdData.Rows = grdData.FixedRows Then Exit Sub
    
        If MsgBox(LoadResString(201), vbQuestion + vbYesNo) = vbYes Then
            m_sFlag = ID_DELETE
        
            Set oCustom = New PlusLib2.CCustom
            oCustom.Connection = adoCon
'            oCustom.UserName = g_sUserName
            If oCustom.DeleteCustom(txtCustomID) Then Call txtSearch_Change

            Set oCustom = Nothing
        End If
    '---------------------------------------------------------------------------
    ElseIf Index = ID_SAVE Then '[4] ����
        If SaveData Then
            Call ChangeMode(Me, True)
            Call SetGrid(FL_BY_NAME)
            If Len(txtSearch) > 0 Then Call txtSearch_Change
            
            m_sFlag = ""
        Else
            MsgBox LoadResString(151), vbCritical
        End If
        grdData.SetFocus
        
    ElseIf Index = ID_CANCEL Then
        m_sFlag = ""
        If grdData.Rows > 1 Then
            Call ShowData
        Else
            Call ClearData
        End If
        Call ChangeMode(Me, True)
        grdData.SetFocus
    End If
    
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "Custom.cmdOperate_Click", Err.Description)
End Sub

Private Sub SetGrid(ByVal Index As EFindClss, Optional sNewData)
    Dim oCustom As PlusLib2.CCustom
    Dim rs As ADODB.Recordset, iLoop As Integer
    Dim lNowRow As Long

    Dim lsAdditemStr As String
    
    On Error GoTo ErrHandler
    
    m_bSkip = True
    
    Set oCustom = New PlusLib2.CCustom
    oCustom.Connection = adoCon
    
    If Index = FL_BY_CODE Then
        If LenB(StrConv(sNewData, vbFromUnicode)) < 4 Then
            Set rs = oCustom.GetCustom(CStr(sNewData))
        Else
            Set oCustom = Nothing
            Exit Sub
        End If
    ElseIf Index = FL_BY_NAME Then
        If IsMissing(sNewData) Then
            sNewData = ""
        Else
            sNewData = Trim(sNewData)
        End If
        Set rs = oCustom.GetCustom(CStr(sNewData))
    End If
    Set oCustom = Nothing
    
    With grdData
        .Redraw = flexRDNone
        If .Rows > .FixedRows Then
            If m_sFlag = ID_ADDNEW Then
                lNowRow = .Rows
            Else
                lNowRow = .Row
            End If
            .Rows = .FixedRows
        Else
            lNowRow = 1
        End If
        
        Do While Not rs.EOF
            iLoop = iLoop + 1
''            'S_201312_���_99 �� ���� ����-OLD�ҽ�
''            .AddItem CStr(iLoop) & vbTab & rs!CustomID & vbTab & rs!KCustom & vbTab & _
''                CheckNull(rs!Phone1) & vbTab & CheckNull(rs!Phone2) & vbTab & _
''                CheckNull(rs!Chief) & vbTab & CheckNull(rs!FaxNO) & vbTab & _
''                CheckNull(rs!CustomNo) & vbTab & CheckNull(rs!Condition) & vbTab & _
''                CheckNull(rs!Category) & vbTab & _
''                CheckNull(rs!Address1) & vbTab & _
''                CheckNull(rs!Address2) & vbTab & CheckNull(rs!ZipCode) & vbTab & _
''                CheckNull(rs!Email) & vbTab & CheckNull(rs!Homepage) & vbTab & _
''                CheckNull(rs!Name) & vbTab & CheckNull(rs!Phone) & vbTab & _
''                CheckNull(rs!TradeID) & vbTab & _
''                CheckNull(rs!UserID) & vbTab & CheckNull(rs!UserPassword) & vbTab & _
''                CheckNull(rs!ECustom) & vbTab & CheckNull(rs!ShortCustom) & vbTab & CheckNull(rs!LossClss) & vbTab & _
''                CheckNull(rs!SpendingClss) & vbTab & CheckNull(rs!workingClss) & vbTab & _
''                CheckNull(rs!CalClss) & vbTab & CheckNull(rs!PointClss)
            
            'S_201312_���_99 �� ���� ����-NEW�ҽ�
            lsAdditemStr = CStr(iLoop)                                                                                      '0)Row ��
            lsAdditemStr = lsAdditemStr & vbTab & rs!CustomID                                                               '1)�ڵ�
            lsAdditemStr = lsAdditemStr & vbTab & rs!KCustom                                                                '2)��ȣ
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Phone1)                                                      '3)��ǥ��ȭ
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Phone2)                                                      '4)��ȭ��ȣ
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Chief)                                                       '5)��ǥ��
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!FaxNO)                                                       '6)�ѽ���ȣ
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!CustomNO)                                                    '7)����ڹ�ȣ
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Condition)                                                   '8)����
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Category)                                                    '9)����
            'S_201312_���_99 �� ���� ����-Address1=>AddressJiBun1 �κ���
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!AddressJiBun1)                                               '10)�����ּ�(1)
            'S_201312_���_99 �� ���� ����-Address2=>AddressJiBun2 �κ��� ����
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!AddressJiBun2)                                               '11)�����ּ�(2)
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!ZipCode)                                                     '12)������ȣ
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Email)                                                       '13)���ڿ���
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Homepage)                                                    '14)Ȩ������
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Name)                                                        '15)����ڸ�
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Phone)                                                       '16)�����ȭ
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!TradeID)                                                     '17)�ŷ�����
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!UserID)                                                      '18)���α��ο�-�ŷ�óID
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!UserPassword)                                                '19)���α��ο�-�ŷ�ópwd
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!ECustom)                                                     '20)�ŷ�ó(����)
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!ShortCustom)                                                 '21)�ŷ�ó(��Ī)
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!LossClss)                                                    '22)����/Loss
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!SpendingClss)                                                '23)�ҿ䷮ ����
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!workingClss)                                                 '24)������ ����
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!CalClss)                                                     '25)ȯ�걸��
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!PointClss)                                                   '26)�Ҽ��� ����
            'S_201312_���_99 �� ���� �߰�-----------------------------------------
            lsAdditemStr = lsAdditemStr & vbTab & ""                                                                        '27)���� - ���� ���߱� ���� �߰�
            lsAdditemStr = lsAdditemStr & vbTab & ""                                                                        '28)���� - ���� ���߱� ���� �߰�
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!OldNNewClss)                                                 '29)�ּұ���
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!GunMoolMngNo)                                                '30)�ǹ�������ȣ
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Address1)                                                    '31)���θ��ּ�1
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!Address2)                                                    '32)���θ��ּ�2
            lsAdditemStr = lsAdditemStr & vbTab & CheckNull(rs!AddressAssist)                                               '33)���θ� ���� �ּ�
            '---------------------------------------------------------------------
                
            .AddItem lsAdditemStr
           
            If (iLoop Mod 2) = 0 Then '// ¦���� ���� �ٲ��ֱ�
                .Row = iLoop
            
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = &HE0E0E0    '&HC0C0C0
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        Call ChangeScroll
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            
            .Col = 20
            .Row = 1
            .RowSel = .Rows - 1
            .CellBackColor = &H80FFFF    '&HC0FFFF
            
            .Row = IIf(.Rows > lNowRow, lNowRow, .Rows - 1)

            .Col = .FixedCols
            .ColSel = .Cols - 1
            
            lblCount.Caption = "�˻��Ǽ� : " & grdData.Rows - 1 & " ��"
            Call ShowData
        Else
            .HighLight = flexHighlightNever
            lblCount.Caption = "�˻��Ǽ� :"
            
            Call ClearData
        End If
        .Redraw = flexRDDirect
    End With
    m_bSkip = False
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oCustom = Nothing
    
    Call ErrorBox(Err.Number, "frmCustom.SetGrid", Err.Description)
End Sub

Private Sub ClearData()
    txtCustomID = ""
    txtKCustom = ""
    txtECustom = ""
    txtShortCustom = ""
    txtCondition = ""
    txtCategory = ""
    txtUserID = ""
    txtUserPassword = ""
    txtChief = ""
    mskCustomNO = ""
    
    cboLoss.ListIndex = 0
    cboSpending.ListIndex = 0
    cboWorking.ListIndex = 0
    cboCalc.ListIndex = 0
    cboPoint.ListIndex = 0
    cboTrade.ListIndex = 0
    
    txtPhone1 = ""
    txtPhone2 = ""
    txtFaxNO = ""
    txtName = ""
    txtPhone = ""
    'S_201312_���_99 �� ���� �߰�---------------------------------------
    optOldNNew(0).Value = True     '���θ��ּҼ���
    txtGunMoolMngNo.Text = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtAddressAssist.Text = ""
    '--------------------------------------------------------------------
    'S_201312_���_99 �� ���� ����(OLD:txtAddress1)
    txtAddressJiBun1.Text = ""
    'S_201312_���_99 �� ���� ����(OLD:txtAddress2)
    txtAddressJiBun2.Text = ""
    
    mskZipCode = ""
    
    txtHomepage = "http://www."
    txtEMail = ""
    
End Sub

'Private Sub SetGrid()
'    Dim iLoop As Integer
'
'    On Error GoTo ErrHandler
'
'    With grdData
'        .Redraw = flexRDNone
'        Select Case m_sFlag
'            Case ID_ADDNEW, ID_UPDATE
'                Call FillGrid
'            Case ID_DELETE
'                If .Rows = 2 Then
'                    .Rows = 1
'                    .HighLight = flexHighlightNever
'
'                    Call ClearData
'                Else
'                    .RemoveItem .Row
'
'                    For iLoop = 1 To .Rows - 1
'                        .TextMatrix(iLoop, 0) = iLoop
'                    Next iLoop
'
'                    Call ChangeScroll
'                    Call ShowData
'                End If
'        End Select
'
'        m_sFlag = ""
'        .Redraw = flexRDDirect
'    End With
'    Exit Sub
'
'ErrHandler:
'    Call ErrorBox(Err.Number, "Custom.SetGrid", Err.Description)
'End Sub

Private Sub ChangeScroll()
    Dim lRows As Long

    On Error GoTo ErrHandler
    
    lRows = GetVisibleVSGridRowCount(grdData)

    With grdData
        .Redraw = flexRDNone
        If .Rows > LIMIT_ROW Then
            .ColWidth(2) = LIMIT_WIDTH - 240
        Else
            .ColWidth(2) = LIMIT_WIDTH
        End If
        .Redraw = flexRDDirect
    End With
    
    If lRows = 0 Then
        cmdOperate(ID_UPDATE).Enabled = False
        cmdOperate(ID_DELETE).Enabled = False
    Else
        cmdOperate(ID_UPDATE).Enabled = True
        cmdOperate(ID_DELETE).Enabled = True
    End If
    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "Custom.ChangeScroll", Err.Description)

End Sub

Private Sub cmdSelect_Click()
    If grdData.Rows = grdData.FixedRows Then Exit Sub
    
    Call SelectData
End Sub

Private Sub Form_Load()
    m_sFlag = ID_CANCEL
    
    Call SetOperate(Me)
    Call InitGrid
    Call MakeCombo
        
    cmdFind.Picture = LoadResPicture("FIND", vbResIcon)
    cmdAll.Picture = LoadResPicture("ALL", vbResIcon)
        
    'S_201312_���_99 �� ���� �߰�
    '��ȣȭ ���� XOR ����� ������ �ʱ�ȭ
''    arrEncCode = Array(1, 84, 62, 23, 59, 48, 66, 11, 43, 93, 37, 50, 43, 19, 77, 29, 5, 69, 49, 21)
    Call SetXorData
End Sub


Private Sub MakeCombo()
    Dim oCode As PlusLib2.CCode
    Dim rs As ADODB.Recordset

    Set oCode = New PlusLib2.CCode
    oCode.Connection = adoCon
    oCode.CodeType = CD_TRADE
    Set rs = oCode.GetCode()
    
    With cboTrade ' �ŷ�����
        Do While Not rs.EOF
            .AddItem rs!Trade
            .ItemData(.NewIndex) = CLng(rs!TradeID)
        
            .ListIndex = 0
            
            rs.MoveNext
        Loop
        rs.Close
    End With
    Set rs = Nothing
    Set oCode = Nothing

    With cboLoss    ' ����/ Loss ���Կ���
        .AddItem "1.����, Loss ����"
        .ItemData(0) = 1
        .AddItem "2.����, Loss ������"
        .ItemData(1) = 2
        .ListIndex = 0
    End With
        
    With cboSpending    ' �ҿ䷮ ������
        .AddItem "1.����� ����"
        .ItemData(0) = 1
        .AddItem "2.Order�� ����"
        .ItemData(1) = 2
        
        .ListIndex = 0
    End With
        
    With cboWorking     ' ������ ������
        .AddItem "1.����� ����"
        .ItemData(0) = 1
        .AddItem "2.Order�� ����"
        .ItemData(1) = 2
    
        .ListIndex = 0
    End With
        
    With cboCalc        ' Meter->Yard ������
        .AddItem "1.Meter / 0.9144"
        .ItemData(0) = 1
        .AddItem "2.Meter * 1.0936"
        .ItemData(1) = 2
    
        .ListIndex = 0
    End With
        
    With cboPoint       ' �Ҽ��� �������
        .AddItem "1.�ݿø�"
        .ItemData(0) = 1
        .AddItem "2.�ø�"
        .ItemData(1) = 2
        .AddItem "3.����"
        .ItemData(2) = 3
        
        .ListIndex = 0
    End With

End Sub

Private Sub InitGrid()
    Call SetVSFlexGrid(grdData)
    With grdData
        .Redraw = flexRDNone
        .Cols = 34                        'S_201312_���_99 �� ���� ���� (OLD:27)
    
        .TextMatrix(0, 0) = "":                 .ColWidth(0) = 0
        .TextMatrix(0, 1) = "�ڵ�":              .ColWidth(1) = 0
        .TextMatrix(0, 2) = "��ȣ":            .ColWidth(2) = LIMIT_WIDTH:  .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "��ǥ��ȭ":        .ColWidth(3) = 1200:         .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(0, 4) = "��ȭ��ȣ":        .ColWidth(4) = 1230:         .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(0, 5) = "��ǥ��":          .ColWidth(5) = 900:          .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(0, 6) = "�ѽ���ȣ":        .ColWidth(6) = 1230:         .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(0, 7) = "����ڹ�ȣ":      .ColWidth(7) = 0
        .TextMatrix(0, 8) = "����":            .ColWidth(8) = 1380:
        .TextMatrix(0, 9) = "����":            .ColWidth(9) = 1380:
        .TextMatrix(0, 10) = "�ּ�(1)":        .ColWidth(10) = 0
        .TextMatrix(0, 11) = "�ּ�(2)":        .ColWidth(11) = 0 '3008
        .TextMatrix(0, 12) = "������ȣ":       .ColWidth(12) = 0
        .TextMatrix(0, 13) = "���ڿ���":       .ColWidth(13) = 0
        .TextMatrix(0, 14) = "Ȩ������":       .ColWidth(14) = 0
        .TextMatrix(0, 15) = "����ڸ�":       .ColWidth(15) = 905:        .ColAlignment(15) = flexAlignCenterCenter
        .TextMatrix(0, 16) = "�����ȭ":       .ColWidth(16) = 1270:       .ColAlignment(16) = flexAlignLeftCenter
        .TextMatrix(0, 17) = "�ŷ�����":       .ColWidth(17) = 0
        .TextMatrix(0, 18) = "�ŷ�óID":       .ColWidth(18) = 0
        .TextMatrix(0, 19) = "�ŷ�ópwd":      .ColWidth(19) = 0
        .TextMatrix(0, 20) = "�ŷ�ó(����)":   .ColWidth(20) = 0
        .TextMatrix(0, 21) = "�ŷ�ó(��Ī)":   .ColWidth(21) = 0
        .TextMatrix(0, 22) = "����/Loss":      .ColWidth(22) = 0
        .TextMatrix(0, 23) = "�ҿ䷮ ����":    .ColWidth(23) = 0
        .TextMatrix(0, 24) = "������ ����":    .ColWidth(24) = 0
        .TextMatrix(0, 25) = "ȯ�걸��":        .ColWidth(25) = 0
        .TextMatrix(0, 26) = "�Ҽ��� ����":     .ColWidth(26) = 0

        'S_201312_���_99 �� ���� �߰�-----------------------------------------------
        .TextMatrix(0, 27) = "":               .ColWidth(27) = 0               '���� ���߱� ���� �߰�
        .TextMatrix(0, 28) = "":               .ColWidth(28) = 0               '���� ���߱� ���� �߰�
        .TextMatrix(0, 29) = "�ּұ���":       .ColWidth(29) = 0
        .TextMatrix(0, 30) = "�ǹ�������ȣ":     .ColWidth(30) = 0
        .TextMatrix(0, 31) = "���θ��ּұ⺻":     .ColWidth(31) = 0
        .TextMatrix(0, 32) = "���θ��ּһ�":     .ColWidth(32) = 0
        .TextMatrix(0, 33) = "���θ������ּ�":     .ColWidth(33) = 0
        
        '//�� ����ColKey ����
        .ColKey(0) = "Idx"
        .ColKey(1) = "CustomID"
        .ColKey(2) = "kCustom"
        .ColKey(3) = "Phone1"
        .ColKey(4) = "Phone2"
        .ColKey(5) = "Chief"
        .ColKey(6) = "FaxNO"
        .ColKey(7) = "CustomNO"
        .ColKey(8) = "Condition"
        .ColKey(9) = "Category"
        .ColKey(10) = "AddressJiBun1"
        .ColKey(11) = "AddressJiBun2"
        .ColKey(12) = "ZipCode"
        .ColKey(13) = "Email"
        .ColKey(14) = "Homepage"
        .ColKey(15) = "Name"
        .ColKey(16) = "Phone"
        .ColKey(17) = "TradeID"
        .ColKey(18) = "UserID"
        .ColKey(19) = "UserPassword"
        .ColKey(20) = "ECustom"
        .ColKey(21) = "ShortCustom"
        .ColKey(22) = "LossClss"
        .ColKey(23) = "SpendingClss"
        .ColKey(24) = "workingClss"
        .ColKey(25) = "CalClss"
        .ColKey(26) = "PointClss"
        .ColKey(27) = "Blank1"
        .ColKey(28) = "Blank2"
        .ColKey(29) = "OldNNewClss"
        .ColKey(30) = "GunMoolMngNo"
        .ColKey(31) = "Address1"
        .ColKey(32) = "Address2"
        .ColKey(33) = "AddressAssist"
        '-----------------------------------------------------------------------
        
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub grdData_AfterSort(ByVal Col As Long, Order As Integer)
    Call ShowData
End Sub

Private Sub grdData_DblClick()
    With grdData
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub
    End With
    
    Call SelectData
End Sub

Private Sub grdData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call grdData_DblClick
    End If
End Sub

Private Sub grdData_RowColChange()
    If m_bSkip Then Exit Sub

    Call ShowData
End Sub

'****************************************************************
'*Author: 2000-06-12 (MON)
'*
'*Description: ��ȸ
'*  �׸��� ���ý� �ش� ������ �ؽ�Ʈ�� ���÷���
'*
'****************************************************************
Private Sub ShowData()
    
    On Error Resume Next
    
    With grdData
    
''        'S_201312_���_99 �� ���� ����-OLD�ҽ�
''        txtCustomID = .TextMatrix(.Row, 1)
''        txtKCustom = .TextMatrix(.Row, 2)
''        txtECustom = .TextMatrix(.Row, 20)
''        txtShortCustom = .TextMatrix(.Row, 21)
''        txtCondition = .TextMatrix(.Row, 8)
''        txtCategory = .TextMatrix(.Row, 9)
''        txtUserID = .TextMatrix(.Row, 18)
''        txtUserPassword = .TextMatrix(.Row, 19)
''        txtChief = .TextMatrix(.Row, 5)
''        mskCustomNO = .TextMatrix(.Row, 7)
''
''        cboLoss.ListIndex = FindComboBox(cboLoss, CLng(.TextMatrix(.Row, 22)))
''        cboSpending.ListIndex = FindComboBox(cboSpending, CLng(.TextMatrix(.Row, 23)))
''        cboWorking.ListIndex = FindComboBox(cboWorking, CLng(.TextMatrix(.Row, 24)))
''        cboCalc.ListIndex = FindComboBox(cboCalc, CLng(.TextMatrix(.Row, 25)))
''        cboPoint.ListIndex = FindComboBox(cboPoint, CLng(.TextMatrix(.Row, 26)))
''        cboTrade.ListIndex = FindComboBox(cboTrade, CLng(.TextMatrix(.Row, 17)))
''
''        txtPhone1 = .TextMatrix(.Row, 3)
''        txtPhone2 = .TextMatrix(.Row, 4)
''        txtFaxNO = .TextMatrix(.Row, 6)
''
''        txtName = .TextMatrix(.Row, 15)
''        txtPhone = .TextMatrix(.Row, 16)
''        txtAddress1 = .TextMatrix(.Row, 10)
''        txtAddress2 = .TextMatrix(.Row, 11)
''        mskZipCode = .TextMatrix(.Row, 12)
''
''        txtHomepage = .TextMatrix(.Row, 14)
''        txtEMail = .TextMatrix(.Row, 13)

        'S_201312_���_99 �� ���� ����-NEW�ҽ�
        txtCustomID = .TextMatrix(.Row, .ColIndex("CustomID"))                                  '�ŷ�ó �ڵ�(1)
        txtKCustom = .TextMatrix(.Row, .ColIndex("kCustom"))                                    '��ȣ(2)
        txtECustom = .TextMatrix(.Row, .ColIndex("ECustom"))                                    '�ŷ�ó(����)(20)
        txtShortCustom = .TextMatrix(.Row, .ColIndex("ShortCustom"))                            '�ŷ�ó(��Ī)(21)
        txtCondition = .TextMatrix(.Row, .ColIndex("Condition"))                                '����(8)
        txtCategory = .TextMatrix(.Row, .ColIndex("Category"))                                  '����(9)
        txtUserID = .TextMatrix(.Row, .ColIndex("UserID"))                                      '���α��ο�-�ŷ�óID(18)
        txtUserPassword = .TextMatrix(.Row, .ColIndex("UserPassword"))                          '���α��ο�-�ŷ�ópwd(19)
        txtChief = .TextMatrix(.Row, .ColIndex("Chief"))                                        '��ǥ��(5)                                       '
        mskCustomNO = .TextMatrix(.Row, .ColIndex("CustomNO"))                                  '����ڹ�ȣ(7)
        cboLoss.ListIndex = FindComboBox(cboLoss, CLng(.TextMatrix(.Row, .ColIndex("LossClss"))))       '����/Loss(22)
        cboSpending.ListIndex = FindComboBox(cboSpending, CLng(.TextMatrix(.Row, .ColIndex("SpendingClss"))))      '�ҿ䷮ ����(23)
        cboWorking.ListIndex = FindComboBox(cboWorking, CLng(.TextMatrix(.Row, .ColIndex("workingClss"))))        '������ ����(24)
        cboCalc.ListIndex = FindComboBox(cboCalc, CLng(.TextMatrix(.Row, .ColIndex("CalClss"))))              'ȯ�걸��(25)
        cboPoint.ListIndex = FindComboBox(cboPoint, CLng(.TextMatrix(.Row, .ColIndex("PointClss"))))            '�Ҽ��� ����(26)
        cboTrade.ListIndex = FindComboBox(cboTrade, CLng(.TextMatrix(.Row, .ColIndex("TradeID"))))            '�ŷ�����(17)
        txtPhone1 = .TextMatrix(.Row, .ColIndex("Phone1"))                                      '��ǥ��ȭ(3)
        txtPhone2 = .TextMatrix(.Row, .ColIndex("Phone2"))                                      '��ȭ��ȣ(4)
        txtFaxNO = .TextMatrix(.Row, .ColIndex("FaxNO"))                                        '�ѽ���ȣ(6)
        txtName = .TextMatrix(.Row, .ColIndex("Name"))                                          '����ڸ�(15)
        txtPhone = .TextMatrix(.Row, .ColIndex("Phone"))                                        '�������ȭ(16)
        mskZipCode = .TextMatrix(.Row, .ColIndex("ZipCode"))                                    '������ȣ(12)
''        txtAddress1 = .TextMatrix(.Row, .ColIndex("AddressJiBun1"))                           '�����ּ�1(10)
''        txtAddress2 = .TextMatrix(.Row, .ColIndex("AddressJiBun2"))                           '�����ּ�2(11)
        
        'S_201312_���_99 �� ���� �߰�-----------------------------------------------------------------
        If .TextMatrix(.Row, .ColIndex("OldNNewClss")) = "0" Then                                     '�ּұ���(29)
            optOldNNew(0).Value = True     '���θ��ּҼ���
        Else
            optOldNNew(1).Value = True     '�����ּ�
        End If
        
        txtGunMoolMngNo.Text = .TextMatrix(.Row, .ColIndex("GunMoolMngNo"))                 '�ǹ����� �����ĺ���ȣ(30)
        txtAddress1.Text = .TextMatrix(.Row, .ColIndex("Address1"))                         ' �ּ�-���θ�(31)
        txtAddress2.Text = .TextMatrix(.Row, .ColIndex("Address2"))                         '�ּ�2-���θ�(32)
        txtAddressAssist.Text = .TextMatrix(.Row, .ColIndex("AddressAssist"))               '���θ� �����ּ�(33)
        '------------------------------------------------------------------------------------------------
        'S_201312_���_99 �� ���� ����(OLD:txtAddress1)
        txtAddressJiBun1.Text = .TextMatrix(.Row, .ColIndex("AddressJiBun1"))               '�����ּ�1(10)
        'S_201312_���_99 �� ���� ����(OLD:txtAddress2)
        txtAddressJiBun2.Text = .TextMatrix(.Row, .ColIndex("AddressJiBun2"))               '�����ּ�2(11)
        txtHomepage = .TextMatrix(.Row, .ColIndex("Homepage"))                              'Ȩ������(14)
        txtEMail = .TextMatrix(.Row, .ColIndex("Email"))                                    '�̸���(13)
        
    End With
    
End Sub


Private Sub mskCustomNO_GotFocus()
    With mskCustomNO
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub mskZipCode_GotFocus()
    With mskZipCode
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

'S_201312_���_99 �� ���� �߰�
Private Sub optOldNNew_Click(Index As Integer)
    If optOldNNew(0).Value = True Then
        fraDoro.Enabled = True
        fraJiBun.Enabled = False
    Else
        fraDoro.Enabled = False
        fraJiBun.Enabled = True
    End If
End Sub

Private Sub optSize_Click(Index As Integer)
    If optSize(0).Value Then    '[0] ���
        grdData.Width = 11820
    Else                        '[1] ��
        grdData.Width = 3495
    End If
End Sub

'S_201312_���_99 �� ���� �߰�
Private Sub txtAddress1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdFind_Click
    End If
End Sub

'S_201312_���_99 �� ���� �߰�
'****************************************************************
'*Description:
'*  ADO�� �̿��Ͽ� ������ �캯��ȣ Database�� �����ϱ�
'****************************************************************
Public Function ConnectWizDB() As Boolean
    
    Dim sWizConnect$

    On Error GoTo ErrHandler

    If adoWizCon Is Nothing Then
        
        If Command() <> "" Then
            '//�׽���
           ' MsgBox "DB Test �ӽ�"
          '  g_sServer = "wizis.iptime.org,1433"
          '  g_sDatabase = "ZipDB"

            If g_sWizSQLAuthType = "1" Then
                
                                'SQL����
                sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & g_sWizSQLID & ";Password=" & g_sWizPassword & _
                            ";Initial Catalog=" & g_sWizDatabase & _
                            ";Data Source=" & g_sWizServer & _
                            ";Use Procedure for Prepare=1;Auto Translate=True;"
                
            Else
                '����������
                sWizConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sWizServer & ";DATABASE=" & g_sWizDatabase & ";UID=sa;PWD=;"
            End If



        Else

            If g_sWizSQLAuthType = "1" Then
                'SQL����
                sWizConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & g_sWizSQLID & ";Password=" & g_sWizPassword & _
                       ";Initial Catalog=" & g_sWizDatabase & _
                       ";Data Source=" & g_sWizServer & _
                       ";Use Procedure for Prepare=1;Auto Translate=True;"
            Else
         
                '����������
                sWizConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sWizServer & ";DATABASE=" & g_sWizDatabase & ";UID=sa;PWD=;"
            End If
        End If
        
        'S_201312_���_99 �� ���� �߰�-������ȣ ��ȸ ���� connection
        Set adoWizCon = New ADODB.Connection
        With adoWizCon
            .CommandTimeout = 60
            .ConnectionString = sWizConnect
            .CursorLocation = adUseClient
            .Open sWizConnect
        End With

        ConnectWizDB = True
    ElseIf adoWizCon.State = adStateOpen Then
        ConnectWizDB = True
    Else
        ConnectWizDB = False
    End If
    
    Exit Function
ErrHandler:
''    Unload frmSplash

''    Call ErrorBox(Err.Number, Err.Source, Err.Description, "DB Connection ����", True)

    ConnectWizDB = False
End Function

'S_201312_���_99 �� ���� �߰�
'================================================================
'*  ���θ� �ּ� �˻��� ���� ������ DB���� ������ �������� ������
'*  ��������: 2013.08.06
'*  ������  : ���¿�
'*  Parameter  :
'*   pUseYN : ��� ����
'*---------------------------------------------------------------
'*  �����̷�:
'*---------------------------------------------------------------
'*  ��������    ������  ���泻��
'================================================================
Public Function Gf_DBConnInfo(pRs As ADODB.Recordset, pUseYN As String) As Boolean
    Dim lssql                           As String
    Dim rs                              As ADODB.Recordset
    On Error GoTo Err_Rtn
    
    lssql = ""
    lssql = lssql & "  SELECT [ConnectioinType]             " & vbCrLf '��������
    lssql = lssql & "       , [SeverCode]                   " & vbCrLf '�����ڵ�
    lssql = lssql & "       , [SeverName]                   " & vbCrLf '������
    lssql = lssql & "       , [SeverAlias]                  " & vbCrLf '������Ī
    lssql = lssql & "       , [SeverAddress]                " & vbCrLf '�����ּ�
    lssql = lssql & "       , [MangCompany]                 " & vbCrLf '������ü
    lssql = lssql & "       , [DBNameMain]                  " & vbCrLf '����DB��
    lssql = lssql & "       , [DBNameSub]                   " & vbCrLf '����DB��
    lssql = lssql & "       , [PortFrom]                    " & vbCrLf '������Ʈ
    lssql = lssql & "       , [PortTo]                      " & vbCrLf '������Ʈ
    lssql = lssql & "       , [AuthCode1]                   " & vbCrLf '�����ڵ�1
    lssql = lssql & "       , [AuthCode2]                   " & vbCrLf '�����ڵ�2
    lssql = lssql & "       , [SQLAuthType]                 " & vbCrLf 'SQL����Ÿ��
    lssql = lssql & "       , [SQLID]                       " & vbCrLf 'SQL�α���ID
    lssql = lssql & "       , [SQLPass]                     " & vbCrLf 'SQL�α��ξ�ȣ
    lssql = lssql & "       , [PassAuthCode]                " & vbCrLf '��ȣ�����ڵ�
    lssql = lssql & "       , [Comments]                    " & vbCrLf 'Comment
    lssql = lssql & "       , [UseYN]                       " & vbCrLf '��뿩��
    lssql = lssql & "    FROM DBConnInfo                    " & vbCrLf
    lssql = lssql & "   WHERE 1= 1                          " & vbCrLf
    
    '��뿩��
    If pUseYN <> "" Then
        lssql = lssql & "   AND  UseYN='" & pUseYN & "'     " & vbCrLf
    End If
    
    lssql = lssql & " order by  [SeverName]                 " & vbCrLf
    
    If Gf_DB_OpenRS(adoCon, pRs, lssql) = False Then GoTo Err_Rtn
    
    Gf_DBConnInfo = True
    
    Exit Function
    
Err_Rtn:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " / " & Err.Description, vbCritical, "[Gf_DBConnInfo]"
    End If
    Call Gs_DB_CloseRs(pRs)
End Function

'S_201312_���_99 �� ���� ����-OLD�ҽ�
''Private Sub cmdFind_Click()
''    Dim oZipFind As PlusFind2.CZipFind
''
''    On Error GoTo ErrHandler
''
''    Set oZipFind = New PlusFind2.CZipFind
''    oZipFind.Connection = adoCon
''''    oZipFind.Address1 = txtName(4)
''    If oZipFind.Show() Then
''        txtAddress1 = oZipFind.Address
''        mskZipCode = oZipFind.ZipCode
''        txtAddress2.SetFocus
''    End If
''    Set oZipFind = Nothing
''    Exit Sub
''ErrHandler:
''    Set oZipFind = Nothing
''
''    Call ErrorBox(Err.Number, "Custom.cmdFind_Click", Err.Description)
''End Sub

'S_201312_���_99 �� ���� ����-NEW�ҽ�
Private Sub cmdFind_Click()
    Dim rs                                  As ADODB.Recordset
    Dim oZipFind As PlusFind2.CZipFind
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    
    'S_201312_���_99 �� ���� �߰�--------------------------------------------------------------------------
    If g_sWizServer = "" Or g_sWizDatabase = "" Then
    
        '-------------------------------------
        ' ���θ� �ּ� �˻��� ���� ������ DB���� ������ �������� ������
        '-------------------------------------
        If Gf_DBConnInfo(rs, "Y") = True Then
        
             If rs.EOF = False Then
    
                'DB���� �о�� DB���� ������ ���α׷������� ��� �ϴ� Global������ �־���
                g_sWizServer = Trim(CheckNull(rs!SeverAddress)) & IIf(Trim(CheckNull(rs!PortFrom)) = "", "", ", " & Trim(CheckNull(rs!PortFrom)))
                g_sWizDatabase = Trim(CheckNull(rs!DBNameMain))
                g_sWizSQLID = Trim(CheckNull(rs!SQLID))
''                g_sWizPassword = Trim(CheckNull(rs!SQLPass))
                
                g_sWizSQLAuthType = Trim(CheckNull(rs!SQLAuthType))         'SQL����Ÿ��(1:SQL,2:������)
                
                'XOR ���� ������ �迭 �缱��
                Call SetXorDataReDim(Len(Trim(CheckNull(rs!PassAuthCode))))
                
                g_sWizPassword = deCode(Trim(CheckNull(rs!PassAuthCode)))          '��ȣȭ �� ���� ��ȣȭ ��
            
''                 g_DBConnInfo.ConnectioinType = Trim(CheckNull(rs!ConnectioinType))  '��������
''                 g_DBConnInfo.SeverCode = Trim(CheckNull(rs!SeverCode))              '�����ڵ�
''                 g_DBConnInfo.SeverName = Trim(CheckNull(rs!SeverName))              '������
''                 g_DBConnInfo.SeverAlias = Trim(CheckNull(rs!SeverAlias))            '������Ī
''                 g_DBConnInfo.SeverAddress = Trim(CheckNull(rs!SeverAddress))        '�����ּ�
''                 g_DBConnInfo.MangCompany = Trim(CheckNull(rs!MangCompany))          '������ü
''                 g_DBConnInfo.DBNameMain = Trim(CheckNull(rs!DBNameMain))            '����DB��
''                 g_DBConnInfo.DBNameSub = Trim(CheckNull(rs!DBNameSub))              '����DB��
''                 g_DBConnInfo.PortFrom = Trim(CheckNull(rs!PortFrom))                '������Ʈ
''                 g_DBConnInfo.PortTo = Trim(CheckNull(rs!PortTo))                    '������Ʈ
''                 g_DBConnInfo.AuthCode1 = Trim(CheckNull(rs!AuthCode1))              '�����ڵ�1
''                 g_DBConnInfo.AuthCode2 = Trim(CheckNull(rs!AuthCode2))              '�����ڵ�2
''                 g_DBConnInfo.SQLAuthType = Trim(CheckNull(rs!SQLAuthType))          'SQL����Ÿ��(1:SQL,2:������)
''                 g_DBConnInfo.SQLID = Trim(CheckNull(rs!SQLID))                      'SQL�α���ID
''                 g_DBConnInfo.SQLPass = Trim(CheckNull(rs!SQLPass))                  'SQL�α��ξ�ȣ
''                 g_DBConnInfo.PassAuthCode = Trim(CheckNull(rs!PassAuthCode))        '��ȣ�����ڵ�
                 
             End If
             
        End If
    
    End If
    
    'S_201312_���_99 �� ���� �߰�
    '������ ������ȣ  DB ���� �����
''    If g_bChkWizDBConn = False Then
''        g_bChkWizDBConn = ConnectWizDB()
''    End If
    If ConnectWizDB() = False Then
        MsgBox "���θ� �ּ� DB���� ���� !!!" & vbCrLf & "�������� ���� ���н� �������� �Է��Ͻʽÿ�.", vbCritical, "DB���� ����"
        Exit Sub
    
    End If
            
    '-------------------------------------------------------------------------------------------------------
    Set oZipFind = New PlusFind2.CZipFind
    
    'S_201312_���_99 �� ���� ����(OLD: adoCon)
    oZipFind.Connection = adoWizCon                            '���θ� �ּҰ��� ������ DB
    
    'S_201312_���_99 �� ���� �߰�
    If optOldNNew(0).Value = True Then      '���θ� �ּ�
        oZipFind.Address1 = txtAddress1
    Else                                    '���� �ּ�
        'S_201312_���_99 �� ���� ����(OLD:oZipFind.Address1,txtAddress1.Text)
        oZipFind.AddressJiBun1 = txtAddressJiBun1.Text
    End If
    
    'S_201312_���_99 �� ���� �߰�
    oZipFind.OldNNewSet = IIf(optOldNNew(0).Value = True, "0", "1")
    
''    oZipFind.Address1 = txtName(4)
    If oZipFind.Show() Then
        'S_201312_���_99 �� ���� ����-----------------------------------------------
        mskZipCode = oZipFind.ZipCode
        If oZipFind.OldNNewClss = "0" Then    '���θ� �ּ�
            optOldNNew(0).Value = True
                
            txtAddress1.Text = oZipFind.Address
            txtAddress2.Text = oZipFind.AddressDetail
            txtAddressAssist.Text = oZipFind.AddressAssist
            txtGunMoolMngNo.Text = oZipFind.GunMoolMngNo

            txtAddress2.SetFocus
        Else
            optOldNNew(1).Value = True
            txtAddressJiBun1.Text = oZipFind.Address
            txtAddressJiBun2.Text = ""                       'S_201312_���_99 �� ���� �߰�
        
            txtAddressJiBun2.SetFocus
        End If
        '----------------------------------------------------------------------------
        
    End If
    Set oZipFind = Nothing
    Exit Sub
ErrHandler:
    Set oZipFind = Nothing
    
    Call ErrorBox(Err.Number, "Custom.cmdFind_Click", Err.Description)
End Sub

'S_201312_���_99 �� ���Ѽ���
Private Sub txtAddressJiBun1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdFind_Click
    End If
End Sub

Private Sub txtKCustom_Change()
    txtShortCustom = txtKCustom
End Sub

Private Sub txtSearch_Change()
    Dim iLoop As Integer, iCols As Integer
    Dim iCount As Integer
    Dim iNowRow As Integer

    If Len(Trim(txtSearch)) > 0 Then
        m_bSkip = True
        With grdData
            .Redraw = flexRDNone
            iCols = .Cols
            
            For iLoop = .FixedRows To .Rows - 1
                If InStr(UCase(.TextArray(iLoop * iCols + 2)), UCase(txtSearch)) = 0 Then
                    .RowHidden(iLoop) = True
                    iCount = iCount + 1
                Else
                    .RowHidden(iLoop) = False
                    iNowRow = iLoop
                End If
            Next iLoop
            
            m_bSkip = False
            If iNowRow > .FixedRows Then
                .Row = iNowRow
                
                .Col = .FixedCols
                .ColSel = .Cols - .FixedCols
            End If
            
            .Redraw = flexRDDirect
            .TopRow = .Row
        End With
    Else
        Call cmdAll_Click
    End If
    
    If iCount > 0 Then
        cmdAll.Visible = True
    Else
        cmdAll.Visible = False
    End If
    
    Call ChangeScroll
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        grdData.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        Call SetGrid(FL_BY_NAME, txtSearch)
    End If
    
End Sub


Public Function SetMsg(SelData(), Optional sNewData) As Boolean
    Dim i%
       
    If IsMissing(sNewData) Then
        Me.Show vbModal
    Else
        If sNewData = "" Then
            Me.Show vbModal
        Else
            Call SetGrid(FL_BY_CODE, sNewData)
            If grdData.Rows = grdData.FixedRows Then
                txtSearch = sNewData
                Call SetGrid(FL_BY_NAME, sNewData)
            End If
            
            '------------------------------------------------
            With grdData
                If .Rows > .FixedRows Then
                    If .Rows = .FixedRows + 1 Then
                        Call SelectData
                    Else
                        Me.Show vbModal
                    End If
                Else
                    If MsgBox(LoadResString(112), vbQuestion + vbYesNo) = vbYes Then
                        Call cmdOperate_Click(ID_ADDNEW)
                        
                        txtKCustom.Text = sNewData
                        txtShortCustom.Text = sNewData
                        Me.Show vbModal
                    Else
                        Me.Show vbModal
                    End If
                End If
            End With
        End If
    End If
    
    '=====================================================================
    If m_bSelected Then
        With grdData
            ReDim SelData(UBound(wData) - 1)
            For i = LBound(wData) To UBound(wData) - 1
                SelData(i) = wData(i)
            Next i
        End With
    End If
    
    SetMsg = m_bSelected
End Function

Private Sub SelectData()
    Dim i%
    
    On Error Resume Next
    
    If grdData.Rows > 1 Then
        m_bSelected = True
        
        ReDim wData(grdData.Cols - 1)
        With grdData
            For i = 1 To .Cols - 1
                wData(i - 1) = .TextMatrix(.Row, i)
            Next i
        End With
        
        Me.Hide
    End If
End Sub

