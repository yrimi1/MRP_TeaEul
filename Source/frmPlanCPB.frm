VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanCPB 
   ClientHeight    =   9285
   ClientLeft      =   -660
   ClientTop       =   2745
   ClientWidth     =   15240
   Icon            =   "frmPlanCPB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15240
   Begin VB.Frame fraData 
      Caption         =   " [  ��ȹ��Ȳ  ]"
      Height          =   4410
      Left            =   30
      TabIndex        =   4
      Top             =   45
      Width           =   15165
      Begin VB.CommandButton cmdCheck 
         Caption         =   "��ü ����"
         Height          =   315
         Index           =   0
         Left            =   1305
         TabIndex        =   28
         Top             =   4050
         Width           =   1140
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "���� ����"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   27
         Top             =   4050
         Width           =   1140
      End
      Begin VB.ComboBox cboProcessID 
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1485
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   25
         Top             =   225
         Width           =   1965
      End
      Begin VSFlex7LCtl.VSFlexGrid grdPlanData 
         Height          =   3420
         Left            =   150
         TabIndex        =   11
         Top             =   600
         Width           =   14955
         _cx             =   26379
         _cy             =   6032
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
      Begin MSComCtl2.DTPicker dtpPlanDate 
         Height          =   360
         Left            =   5625
         TabIndex        =   5
         Top             =   225
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   73269248
         CurrentDate     =   36871
      End
      Begin VSFlex7LCtl.VSFlexGrid grdTotal 
         Height          =   360
         Left            =   10500
         TabIndex        =   24
         Top             =   4005
         Width           =   4560
         _cx             =   8043
         _cy             =   635
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
      Begin Threed.SSPanel pnlCaption 
         Height          =   360
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "����"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdLeft 
         Height          =   390
         Left            =   5040
         TabIndex        =   34
         Top             =   195
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   688
         _Version        =   196609
         AutoSize        =   1
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdRight 
         Height          =   390
         Left            =   8295
         TabIndex        =   35
         Top             =   195
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   688
         _Version        =   196609
         AutoSize        =   1
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   360
         Index           =   1
         Left            =   3705
         TabIndex        =   36
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "��ȹ����"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdBring 
         Height          =   360
         Left            =   8910
         TabIndex        =   37
         Top             =   210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         _Version        =   196609
         Caption         =   "��������"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame fraKey 
      Caption         =   " [  ��ȹ�ۼ�  ]"
      Height          =   1605
      Left            =   30
      TabIndex        =   1
      Top             =   4560
      Width           =   15165
      Begin VB.TextBox txtOrderID 
         Alignment       =   2  '��� ����
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         TabIndex        =   32
         Top             =   570
         Width           =   2175
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  '������ ����
         Height          =   315
         Left            =   9075
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   570
         Width           =   1215
      End
      Begin VB.TextBox txtColorName 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   6375
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   570
         Width           =   2685
      End
      Begin VB.TextBox txtRemark 
         Height          =   615
         Left            =   3660
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   930
         Width           =   6630
      End
      Begin VB.TextBox txtPersonID 
         Alignment       =   2  '��� ����
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   5025
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   570
         Width           =   1335
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&S)"
         Height          =   795
         Index           =   3
         Left            =   11130
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   10
         ToolTipText     =   "�ڷ� ����"
         Top             =   720
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "�߰�(&A)"
         Height          =   795
         Index           =   0
         Left            =   12720
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   9
         ToolTipText     =   "�ڷ� �߰�"
         Top             =   720
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&D)"
         Height          =   795
         Index           =   2
         Left            =   14310
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   8
         ToolTipText     =   "�ڷ� ����"
         Top             =   720
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&U)"
         Height          =   795
         Index           =   1
         Left            =   13515
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   7
         ToolTipText     =   "�ڷ� ����"
         Top             =   720
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "���(&C)"
         Height          =   795
         Index           =   4
         Left            =   11925
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   6
         ToolTipText     =   "�ڷ� ���"
         Top             =   720
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.ComboBox cboEmerClss 
         Height          =   300
         ItemData        =   "frmPlanCPB.frx":000C
         Left            =   3675
         List            =   "frmPlanCPB.frx":000E
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   3
         Top             =   570
         Width           =   1305
      End
      Begin VB.ComboBox cboPlanClss 
         Height          =   300
         ItemData        =   "frmPlanCPB.frx":0010
         Left            =   2325
         List            =   "frmPlanCPB.frx":0012
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   2
         Top             =   570
         Width           =   1305
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   2
         Left            =   3675
         TabIndex        =   12
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
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
         Caption         =   "��ޱ���"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   3
         Left            =   2325
         TabIndex        =   13
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
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
         Caption         =   "��ȹ����"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   4
         Left            =   5025
         TabIndex        =   14
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         Caption         =   "�ۼ���"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   5
         Left            =   2325
         TabIndex        =   15
         Top             =   945
         Width           =   1305
         _ExtentX        =   2302
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
         Caption         =   "����"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlColorName 
         Height          =   300
         Left            =   6375
         TabIndex        =   21
         Top             =   240
         Width           =   2685
         _ExtentX        =   4736
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
         Caption         =   "�����"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   0
         Left            =   9075
         TabIndex        =   23
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "����"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Left            =   135
         TabIndex        =   29
         Top             =   945
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   75
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "���� ��ȣ"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   330
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin Threed.SSPanel chkSearch 
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   33
         Top             =   225
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   196609
         Caption         =   "���� ��ȣ"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13590
      TabIndex        =   0
      Top             =   8595
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      �ݱ�(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   2415
      Left            =   30
      TabIndex        =   16
      Top             =   6180
      Width           =   15180
      _cx             =   26776
      _cy             =   4260
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
      Begin VB.CheckBox chkExpand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����Ȯ��"
         Height          =   500
         Left            =   0
         Style           =   1  '�׷���
         TabIndex        =   19
         Top             =   0
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmPlanCPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------
Private m_nSelected As Integer ' ���� ���ð���
'Private m_bSkipEvent As Boolean
Private m_bLoading As Boolean
Private m_iFlag    As Integer   ' ���� ���� (�߰�/����/����/�˻�)
Private m_ProcessID As String
'---------------------------------------------------------------

Private Sub cboProcessID_Click()
    
''    cboProcessID.AddItem "C����"  '4000   Į��  0
''    cboProcessID.AddItem "����"   '4300   Į����  1
    
    Call InitGrid
    Call ClearData
    txtOrderID.Text = ""
    If cboProcessID.ListIndex = 0 Then    'C����
        m_ProcessID = "4000"
        pnlColorName.Visible = True
        txtColorName.Visible = True
        grdPlanData.ColHidden(7) = False
        grdPlanData.ColHidden(8) = False
        pnlName(0).Visible = True
        txtQty.Visible = True
    Else
        m_ProcessID = "4300"
        pnlColorName.Visible = False
        txtColorName.Visible = False
        pnlName(0).Visible = False
        txtQty.Visible = False
        grdPlanData.ColHidden(7) = True
        grdPlanData.ColHidden(8) = True
    End If
    Call FillGrdPlanCPB
End Sub

Private Sub chkExpand_Click()
    Dim i%
    With grdOrder
        For i = 13 To 23
            .ColHidden(i) = IIf(chkExpand.Value = vbChecked, False, True)
        Next i
        
        .ColHidden(14) = True
        .ColHidden(15) = True
        .ColHidden(17) = True
        .ColHidden(20) = True
        
        If chkExpand.Value Then
            .ScrollBars = flexScrollBarBoth
        Else
            .ScrollBars = flexScrollBarVertical
        End If
    End With
End Sub


''Private Sub cmdExcel_Click()
''    If grdPlanData.Rows = 1 Then
''        MsgBox LoadResString(111), vbInformation
''        cmdSearch.SetFocus
''
''        Exit Sub
''    End If
''    Call MakeExcelGrid(grdPlanData)
''End Sub


Private Sub NonEditMode(ByVal NewValue As Boolean)
    Dim i%

    fraData.Enabled = NewValue
    txtOrderID.Enabled = NewValue
        
'    If NewValue Then '[1] ��ȸ��� = True
'        grdPlanData.Editable = flexEDNone
'    Else '[2] ������� = False
'        grdPlanData.Editable = flexEDKbdMouse
'    End If

    cboEmerClss.Locked = NewValue
    cboPlanClss.Locked = NewValue
    txtRemark.Locked = NewValue
    txtQty.Locked = NewValue
End Sub

Private Sub ClearData()
    cboEmerClss.ListIndex = 0
    cboPlanClss.ListIndex = 0
    txtPersonID.Text = ""
    txtRemark.Text = ""
    txtPersonID.Text = g_sPersonName
    txtPersonID.Tag = g_sUserName
End Sub

Private Sub cmdBring_Click()
    Dim oPlanCPB As PlusLib2.CPlanCPB
    Dim sOrderIDs As String
    Dim i%

    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrHandler
    
    sOrderIDs = ""
    Set oPlanCPB = New PlusLib2.CPlanCPB
    oPlanCPB.Connection = g_adoCon
    
    With grdPlanData
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
                Call oPlanCPB.AddNewPlanCPB_Today(Format$(Now, "yyyymmdd"), MakeDate(DF_SHORT, dtpPlanDate) _
                                     , m_ProcessID, txtOrderID.Text, val(.TextMatrix(i, 10)))
            End If
        Next i
    End With

    dtpPlanDate = Now
    
    Call FillGrdPlanCPB

    Screen.MousePointer = vbDefault
    
    Exit Sub
    '-----------------------------------------------------------------------------------------
ErrHandler:
    Screen.MousePointer = vbDefault
    Set oPlanCPB = Nothing
    
    Call ErrorBox(Err.Number, "oPlanCPB.SaveData", Err.Description)
End Sub

Private Sub cmdLeft_Click()
    dtpPlanDate = dtpPlanDate - 1
    Call dtpPlanDate_Change
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    If MakeDate(DF_SHORT, dtpPlanDate) < g_sysDate Then
        MsgBox ("���� ������ �����ʹ� ���, ����, ������ �Ұ��� �մϴ�.")
        Exit Sub
    End If
    
    Select Case Index
        '-------------------------------------------------------------------------------------'
        Case ID_ADDNEW
            m_iFlag = ID_ADDNEW

            Call ClearData
            Call ChangeMode(Me, False)
            Call NonEditMode(False)
            
            If grdOrder.Rows <= grdOrder.FixedRows Then
                MsgBox "���ֹ�ȣ�� ���� �Է� �Ͻʽÿ�.", vbInformation
                Call SetCancel
            End If
            
            Select Case m_ProcessID
                Case "4000"  'c���� -> Į�� ����.
                    With grdOrder
                        If .Rows > .FixedRows Then
                            If .IsSubtotal(grdOrder.Row) = True Then
                                MsgBox "���� ������ �����Ͻʽÿ�", vbInformation
                                
                                'cancel�� ���� ó��
                                Call SetCancel
                            Else
                                txtColorName.Text = grdOrder.TextMatrix(grdOrder.Row, 3)
                                txtColorName.Tag = GetOrderSeq(txtOrderID, txtColorName)
                            End If
                         End If
                    End With
                Case "4300"  '����  ->
                    txtColorName.Text = ""
                    txtColorName.Tag = ""
            End Select
        
        '-------------------------------------------------------------------------------------'
        Case ID_UPDATE
            '���ڵ尡 ���� ���
            If grdPlanData.Rows = grdPlanData.FixedRows Then
                MsgBox LoadResString(111), vbInformation
                Exit Sub
            End If
            
            Call ShowData(grdPlanData.TextMatrix(grdPlanData.Row, 5), grdPlanData.TextMatrix(grdPlanData.Row, 5), grdPlanData.TextMatrix(grdPlanData.Row, 10))

            m_iFlag = ID_UPDATE
            
            Call ChangeMode(Me, False)
            Call NonEditMode(False)

        '-------------------------------------------------------------------------------------'
        Case ID_DELETE
            If grdPlanData.Rows = grdPlanData.FixedRows Then Exit Sub
            
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo) = vbYes Then
                If DeleteData() Then
                    Call NonEditMode(True)
                    Call FillGrdPlanCPB
                    Call ClearData
                End If
            End If
            
        '-------------------------------------------------------------------------------------'
        Case ID_SAVE
            If Not CheckData() Then Exit Sub
            If SaveData() Then
                Call ChangeMode(Me, True)
                Call NonEditMode(True)
                Call FillGrdPlanCPB
              
                m_iFlag = -1
            End If
            
        '-------------------------------------------------------------------------------------'
        Case ID_CANCEL
            Call SetCancel
''            m_iFlag = -1
''            Call ChangeMode(Me, True)
            Call NonEditMode(True)
    End Select

    Exit Sub
    
ErrHandler:
    Call ErrorBox(Err.Number, "Order.cmdOperate_Click", Err.Description)

End Sub

Sub SetCancel()
    txtColorName.Text = ""
    txtColorName.Tag = ""
    m_iFlag = -1
    Call ChangeMode(Me, True)
    Call NonEditMode(True)
    
End Sub
Function CheckData() As Boolean
    Dim dPersonID As String, dPersonName As String
    
    CheckData = True
    
'    If Len(txtRemark) = 0 Or Len(txtOrderID) = 0 Then
    If Len(txtOrderID) = 0 Then
        CheckData = False
    End If
End Function

Private Sub cmdRight_Click()
    dtpPlanDate = dtpPlanDate + 1
    Call dtpPlanDate_Change
End Sub

Private Sub dtpPlanDate_Change()
    grdOrder.Rows = grdOrder.FixedRows
    txtOrderID.Text = ""
    Call ClearData
    Call FillGrdPlanCPB
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub
Private Sub FillGridOrder()
    Dim oPlanInput As PlusLib2.CPlanInput
    Dim rs As ADODB.Recordset
    Dim i%, nTop%
    Dim nNoPlanQty#, nProceTotalQty#
    
    On Error GoTo ErrHandler
    
    m_bLoading = True
    
    Set oPlanInput = New PlusLib2.CPlanInput
    oPlanInput.Connection = g_adoCon
    
    
'''GetOrder(Optional nChkDate As Integer, Optional sSDate As String, Optional sEDate As String, _
'''                    Optional nChkCustomID As Integer, Optional sCustomID As String, _
'''                    Optional nChkArticleID As Integer, Optional sArticleID As String, _
'''                    Optional nChkOrder As Integer, Optional sOrder As String, _
'''                    Optional nChkCloseClss As Integer, Optional nChkStuffClose As Integer)
    Set rs = oPlanInput.GetOrder(0, "", "", 0, "", 0, "", 1, txtOrderID, 1, 0)
    Set oPlanInput = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        
        grdOrder.Rows = grdOrder.FixedRows
        Exit Sub
    End If
    
    With grdOrder
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            nNoPlanQty = rs!ColorQtyYDS * (1 + rs!ChunkRate / 100) - (rs!InstQty - rs!���Qty + rs!���TQty)    '�̰�ȹ��
            nProceTotalQty = rs!��ó��Qty + rs!ȿ��ȣ��Qty + rs!����FQty + rs!����SQty + rs!����SQty + rs!����Qty + _
                            rs!SKQty + rs!����Qty + rs!PeachQty + rs!C����Qty + rs!����Qty + rs!P����Qty + rs!R����Qty + _
                            rs!����Qty + rs!����Qty + rs!�˻�Qty + rs!PauseQty
                            
            If rs!OrderID <> MakeOrderID(.TextMatrix(nTop, 3), OM_REDUCE) Then
                .AddItem "" & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                    rs!ChunkRate & vbTab & SetCurrency(rs!OrderQty) & IIf(rs!UnitClss = "1", " M", "   ") & vbTab & rs!StuffInQty & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & CheckNull(rs!PatternID) & vbTab & rs!WorkWidth
                
                
                Call DoFlexGridGroup(grdOrder, .Rows - 1, 1)
                Call GridCollapse(grdOrder, nTop)
                
                nTop = .Rows - 1
            End If
        
            .AddItem "" & vbTab & "" & vbTab & "" & vbTab & rs!Color & vbTab & rs!Color & vbTab & _
                "" & vbTab & rs!ColorQty & vbTab & "0" & vbTab & nNoPlanQty & vbTab & _
                rs!InstQty - rs!���Qty + rs!���TQty & vbTab & rs!InstQty - rs!���Qty & vbTab & rs!���TQty & vbTab & nProceTotalQty & vbTab & _
                rs!��ó��Qty + rs!ȿ��ȣ��Qty + rs!����FQty + rs!����SQty + rs!����SQty & vbTab & _
                rs!����Qty & vbTab & rs!SKQty & vbTab & _
                rs!����Qty & vbTab & rs!PeachQty & vbTab & rs!C����Qty & vbTab & _
                rs!����Qty + rs!P����Qty + rs!R����Qty & vbTab & rs!����Qty & vbTab & rs!����Qty & vbTab & _
                rs!�˻�Qty & vbTab & rs!PauseQty & vbTab & rs!PassQty & vbTab & rs!DefectQty & vbTab & rs!OutQty & vbTab & CheckNull(rs!PatternID) & vbTab & rs!WorkWidth
        
'            .TextMatrix(nTop, 7) = CLng(.TextMatrix(nTop, 7)) + rs!StuffInQty
            .TextMatrix(nTop, 8) = CLng(.TextMatrix(nTop, 8)) + nNoPlanQty
            .TextMatrix(nTop, 9) = CLng(.TextMatrix(nTop, 9)) + rs!InstQty - rs!���Qty + rs!���TQty
            .TextMatrix(nTop, 10) = CLng(.TextMatrix(nTop, 10)) + rs!InstQty - rs!���Qty
            .TextMatrix(nTop, 11) = CLng(.TextMatrix(nTop, 11)) + rs!���TQty
            .TextMatrix(nTop, 12) = CLng(.TextMatrix(nTop, 12)) + nProceTotalQty
            .TextMatrix(nTop, 13) = CLng(.TextMatrix(nTop, 13)) + rs!��ó��Qty + rs!ȿ��ȣ��Qty + rs!����FQty + rs!����SQty + rs!����SQty
            .TextMatrix(nTop, 14) = CLng(.TextMatrix(nTop, 14)) + rs!����Qty
            .TextMatrix(nTop, 15) = CLng(.TextMatrix(nTop, 15)) + rs!SKQty
            .TextMatrix(nTop, 16) = CLng(.TextMatrix(nTop, 16)) + rs!����Qty
            .TextMatrix(nTop, 17) = CLng(.TextMatrix(nTop, 17)) + rs!PeachQty
            .TextMatrix(nTop, 18) = CLng(.TextMatrix(nTop, 18)) + rs!C����Qty
            .TextMatrix(nTop, 19) = CLng(.TextMatrix(nTop, 19)) + rs!����Qty + rs!P����Qty + rs!R����Qty
            .TextMatrix(nTop, 20) = CLng(.TextMatrix(nTop, 20)) + rs!����Qty
            .TextMatrix(nTop, 21) = CLng(.TextMatrix(nTop, 21)) + rs!����Qty
            .TextMatrix(nTop, 22) = CLng(.TextMatrix(nTop, 22)) + rs!�˻�Qty
            .TextMatrix(nTop, 23) = CLng(.TextMatrix(nTop, 23)) + rs!PauseQty
            .TextMatrix(nTop, 24) = CLng(.TextMatrix(nTop, 24)) + rs!PassQty
            .TextMatrix(nTop, 25) = CLng(.TextMatrix(nTop, 25)) + rs!DefectQty
            .TextMatrix(nTop, 26) = CLng(.TextMatrix(nTop, 26)) + rs!OutQty
            
            rs.MoveNext
        Next i
        rs.Close
        
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > .FixedRows, .FixedRows, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
            MsgBox LoadResString(203), vbInformation
        End If
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    m_bLoading = False
    Call SetGrdShrink(grdOrder, OM_EXPAND)
    
'    If grdOrder.Rows > grdOrder.FixedRows Then
'        Call GridCollapse(grdOrder, nTop)
'    End If
    Exit Sub

ErrHandler:
    m_bLoading = False
    Set oPlanInput = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmPlanCPB.FillGridOrder", Err.Description)
End Sub

Private Sub Form_Load()
    Dim i%

'    m_bLoading = True

    Me.Move 0, 0, 15360, 9840
    

    Call InitGrid
'    Call chkExpand_Click
    Call SetOperate(Me)
    
    Show
    
'    m_bSkipEvent = True

    With cboEmerClss
        .Clear
        .AddItem "����"
        .AddItem "���"
        .ListIndex = 0
    End With
    
    With cboPlanClss
        .Clear
        .AddItem "����"
        .AddItem "����"
        .ListIndex = 0
    End With

'    m_bLoading = False
    
    Call SetDtpDate(2, dtpPlanDate, dtpPlanDate)
    
    '-- �ʼ��Է� �׸� icon ���� �ϱ�

    pnlName(5).Picture = LoadResPicture("BASIC", vbResIcon)
    cmdLeft.Picture = LoadResPicture("LEFT", vbResIcon)
    cmdRight.Picture = LoadResPicture("RIGHT", vbResIcon)
    
    '�������� �޺��ڽ��� C����(C.P.B����), ����(Rapid����)���� �����ϴ� ���ν��� ȣ��
    'Call SetProcessID(cboProcessID, "'4000', '4300'")
    
    cboProcessID.AddItem "C����"  '4000
    cboProcessID.AddItem "����"   '4300
    cboProcessID.ListIndex = 0
    
'    Call FillGrdPlanCPB
    
    Call NonEditMode(True)
    
    txtOrderID.SetFocus
    
End Sub
Private Sub LoadPlanCPB(ByVal pProcID As Integer)
    cboProcessID.ListIndex = pProcID  '0: 4000(c����), 1:Rapid �������� ����
    txtOrderID.SetFocus

End Sub
Private Function DeleteData() As Boolean
    Dim oPlanCPB As PlusLib2.CPlanCPB
    
    On Error GoTo ErrHandler

    DeleteData = False
    
    Set oPlanCPB = New PlusLib2.CPlanCPB
    oPlanCPB.Connection = g_adoCon
    oPlanCPB.UserName = g_sUserName
    
    DeleteData = oPlanCPB.DeletePlanCPB(MakeDate(DF_SHORT, dtpPlanDate), m_ProcessID, MakeOrderID(grdPlanData.TextMatrix(grdPlanData.Row, 5), OM_REDUCE) _
                                    , grdPlanData.TextMatrix(grdPlanData.Row, 10))
    
    Set oPlanCPB = Nothing
    Exit Function
ErrHandler:
    Set oPlanCPB = Nothing

    Call ErrorBox(Err.Number, "frmPlanCPB.DeleteData", Err.Description)
    
End Function

Private Function SaveData() As Boolean
    Dim nColorRow%, i%

    Dim TPlanCPB As PlusLib2.TPlanCPB
    Dim oPlanCPB As PlusLib2.CPlanCPB
    
    Set oPlanCPB = New PlusLib2.CPlanCPB

    Screen.MousePointer = vbHourglass
    
'    On Error GoTo ErrHandler

    With TPlanCPB
        If m_iFlag = ID_ADDNEW Then
            .sJobFlag = "I"
        Else
            .sJobFlag = "U"
        End If
        
        .sPlanDate = MakeDate(DF_SHORT, dtpPlanDate)          '[2] ��ȹ����
        .sProcessID = IIf(cboProcessID.ListIndex = 0, "4000", "4300")             '[3] �����ڵ�
        .sOrderID = Trim(txtOrderID.Text)                     '[4] ������ȣ
        .sPlanClss = Trim(cboPlanClss.Text)                   '[5] ����, ����
        .sEmerClss = Trim(cboEmerClss.Text)                   '[6] ���, ����
        .sPersonID = g_sUserName                              '[7] �ۼ��� �ڵ�
        .sRemark = IIf(Len(Trim(txtRemark.Text)) = 0, " ", Trim(txtRemark.Text))                     '[8] ��ȹ����
        .nOrderSeq = val(txtColorName.Tag)                    ' colorSeq
        .nQty = val(txtQty)
    End With
    
    '-----------------------------------------------------------------------------------------
    oPlanCPB.Connection = g_adoCon
    
    SaveData = oPlanCPB.AddNewPlanCPB(TPlanCPB)
    
    Set oPlanCPB = Nothing
    Screen.MousePointer = vbDefault
    
    Exit Function
    '-----------------------------------------------------------------------------------------
ErrHandler:
    Screen.MousePointer = vbDefault
    Set oPlanCPB = Nothing
    
    Call ErrorBox(Err.Number, "oPlanCPB.SaveData", Err.Description)
End Function



Private Sub cmdCheck_Click(Index As Integer)
    Call SetGridToggleChecked(grdPlanData, Index)
End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub FillGridOrder22()
    Dim oPlanCPB As PlusLib2.CPlanCPB
    Dim rs As ADODB.Recordset
    Dim i%, nTop%
    Dim nNoPlanQty#, nProceTotalQty#
    
  '  On Error GoTo ErrHandler
    
    
    m_bLoading = True
    
    Set oPlanCPB = New PlusLib2.CPlanCPB
    oPlanCPB.Connection = g_adoCon
    
    Set rs = oPlanCPB.GetCPBOrder(Trim$(txtOrderID))
                                 
    Set oPlanCPB = Nothing
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        
        grdOrder.Rows = grdOrder.FixedRows
        txtOrderID.Text = ""
        txtOrderID.Tag = ""
        Exit Sub
    Else
        Call SetOrderID(rs!OrderID, rs!OrderNo)
    End If
    
    With grdOrder
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 0 To rs.RecordCount - 1
            nNoPlanQty = rs!ColorQtyYDS * (1 + rs!ChunkRate / 100) - (rs!InstQty - rs!���Qty + rs!���TQty)    '�̰�ȹ��
            
            nProceTotalQty = rs!��ó��Qty + rs!ȣ��Qty + rs!����Qty + rs!����FQty + rs!����SQty + rs!����SQty + rs!����Qty + rs!����Qty + _
                            rs!PeachQty + rs!C����Qty + rs!����Qty + rs!P����Qty + rs!R����Qty + rs!����Qty + rs!����Qty + _
                            rs!�˻�Qty + rs!PauseQty
                            
            If rs!OrderID <> MakeOrderID(.TextMatrix(nTop, 3), OM_REDUCE) Then
                .AddItem "" & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                    rs!ChunkRate & vbTab & rs!OrderQty & vbTab & rs!StuffInQty & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & _
                    0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & CheckNull(rs!PatternID)
                
                
                Call DoFlexGridGroup(grdOrder, .Rows - 1, 1)
                Call GridCollapse(grdOrder, nTop)
                
                nTop = .Rows - 1
            End If
        
            .AddItem "" & vbTab & "" & vbTab & "" & vbTab & rs!Color & vbTab & rs!Color & vbTab & _
                "" & vbTab & rs!ColorQty & vbTab & 0 & vbTab & nNoPlanQty & vbTab & _
                rs!InstQty - rs!���Qty + rs!���TQty & vbTab & rs!InstQty - rs!���Qty & vbTab & rs!���TQty & vbTab & nProceTotalQty & vbTab & _
                rs!��ó��Qty + rs!ȣ��Qty + rs!����Qty + rs!����FQty + rs!����SQty + rs!����SQty & vbTab & _
                rs!����Qty + rs!����Qty & vbTab & rs!PeachQty & vbTab & rs!C����Qty & vbTab & _
                rs!����Qty + rs!P����Qty + rs!R����Qty & vbTab & rs!����Qty & vbTab & rs!����Qty & vbTab & _
                rs!�˻�Qty & vbTab & rs!PauseQty & vbTab & rs!PassQty & vbTab & rs!DefectQty & vbTab & rs!OutQty & vbTab & CheckNull(rs!PatternID)
        
'            .TextMatrix(nTop, 7) = CLng(.TextMatrix(nTop, 7)) + rs!StuffInQty
            .TextMatrix(nTop, 8) = CLng(.TextMatrix(nTop, 8)) + nNoPlanQty
            .TextMatrix(nTop, 9) = CLng(.TextMatrix(nTop, 9)) + rs!InstQty - rs!���Qty + rs!���TQty
            .TextMatrix(nTop, 10) = CLng(.TextMatrix(nTop, 10)) + rs!InstQty - rs!���Qty
            .TextMatrix(nTop, 11) = CLng(.TextMatrix(nTop, 11)) + rs!���TQty
            .TextMatrix(nTop, 12) = CLng(.TextMatrix(nTop, 12)) + nProceTotalQty
            .TextMatrix(nTop, 13) = CLng(.TextMatrix(nTop, 13)) + rs!��ó��Qty + rs!ȣ��Qty + rs!����Qty + rs!����FQty + rs!����SQty + rs!����SQty
            .TextMatrix(nTop, 14) = CLng(.TextMatrix(nTop, 14)) + rs!����Qty + rs!����Qty
            .TextMatrix(nTop, 15) = CLng(.TextMatrix(nTop, 15)) + rs!PeachQty
            .TextMatrix(nTop, 16) = CLng(.TextMatrix(nTop, 16)) + rs!C����Qty
            .TextMatrix(nTop, 17) = CLng(.TextMatrix(nTop, 17)) + rs!����Qty + rs!P����Qty + rs!R����Qty
            .TextMatrix(nTop, 18) = CLng(.TextMatrix(nTop, 18)) + rs!����Qty
            .TextMatrix(nTop, 19) = CLng(.TextMatrix(nTop, 19)) + rs!����Qty
            .TextMatrix(nTop, 20) = CLng(.TextMatrix(nTop, 20)) + rs!�˻�Qty
            .TextMatrix(nTop, 21) = CLng(.TextMatrix(nTop, 21)) + rs!PauseQty
            .TextMatrix(nTop, 22) = CLng(.TextMatrix(nTop, 22)) + rs!PassQty
            .TextMatrix(nTop, 23) = CLng(.TextMatrix(nTop, 23)) + rs!DefectQty
            .TextMatrix(nTop, 24) = CLng(.TextMatrix(nTop, 24)) + rs!OutQty
            
            
            rs.MoveNext
        Next i
        rs.Close
        
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = IIf(.Rows > .FixedRows, .FixedRows, .Rows - 1)
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
        Else
            .HighLight = flexHighlightNever
        End If
        
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    m_bLoading = False
    Exit Sub

ErrHandler:
    m_bLoading = False
    Set oPlanCPB = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmPlanCPB.FillGridOrder", Err.Description)
End Sub
Private Sub GridCollapse(oFlex As VSFlexGrid, Row As Integer)
    With oFlex
        If Row < .FixedRows Then Exit Sub

        If .IsCollapsed(Row) = flexOutlineCollapsed Then
            .IsCollapsed(Row) = flexOutlineExpanded
        Else
            .IsCollapsed(Row) = flexOutlineCollapsed
        End If
    End With
End Sub
Private Sub DoFlexGridGroup(oFlex As VSFlexGrid, iRow As Integer, iLvl As Integer)
    With oFlex
        ' Set the row as a group
        .IsSubtotal(iRow) = True
        ' Set the indentation level of the group
        .RowOutlineLevel(iRow) = iLvl

        Select Case iLvl
        Case 0
            .Cell(flexcpForeColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
        Case 1, 2
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HFFFFC0    '&HE0E0E0
        End Select
    End With
End Sub

Private Sub InitGrid()
    Dim i%, nWidth&

    '������Ȳ�� ������Ȳ
    With grdOrder
        .Cols = 29

        .Redraw = flexRDNone

        .Rows = 2
        .FixedRows = 2
        .FixedCols = 0
        .FrozenCols = 5
        
        .RowHeight(0) = 250
        .RowHeight(1) = 250

        .TextArray(0) = " ":            .ColWidth(0) = 500
        .TextArray(1) = "�ŷ�ó":       .ColWidth(1) = 1750:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "ǰ��":         .ColWidth(2) = 1550:             .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "������ȣ" & vbCrLf & "��  ��  ��":     .ColWidth(3) = 1350:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "Order No." & vbCrLf & "��  ��  ��":   .ColWidth(4) = 0:               .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "����":         .ColWidth(5) = 800:            .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "���ַ�":       .ColWidth(6) = 1000:            .ColAlignment(6) = flexAlignRightCenter
        .TextArray(7) = "�԰�":       .ColWidth(7) = 900:             .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "�̰�ȹ��":     .ColWidth(8) = 900:             .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "��ȹ��":       .ColWidth(9) = 900:             .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "���":        .ColWidth(10) = 900:            .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "���":        .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignRightCenter
        .TextArray(12) = "������":      .ColWidth(12) = 900:            .ColAlignment(12) = flexAlignRightCenter
        .TextArray(13) = "����":        .ColWidth(13) = 900:            .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "����":        .ColWidth(14) = 900:            .ColAlignment(14) = flexAlignRightCenter
        .TextArray(15) = "S/K":         .ColWidth(15) = 900:            .ColAlignment(15) = flexAlignRightCenter
        .TextArray(16) = "SETT":        .ColWidth(16) = 900:            .ColAlignment(16) = flexAlignRightCenter
        .TextArray(17) = "PEACH":       .ColWidth(17) = 900:            .ColAlignment(17) = flexAlignRightCenter
        .TextArray(18) = "CPB":         .ColWidth(18) = 900:            .ColAlignment(18) = flexAlignRightCenter
        .TextArray(19) = "����":        .ColWidth(19) = 900:            .ColAlignment(19) = flexAlignRightCenter
        .TextArray(20) = "DRY":         .ColWidth(20) = 900:            .ColAlignment(20) = flexAlignRightCenter
        .TextArray(21) = "����":        .ColWidth(21) = 900:            .ColAlignment(21) = flexAlignRightCenter
        .TextArray(22) = "�˻�":        .ColWidth(22) = 900:            .ColAlignment(22) = flexAlignRightCenter
        .TextArray(23) = "����":        .ColWidth(23) = 900:            .ColAlignment(23) = flexAlignRightCenter
        .TextArray(24) = "�˻�":        .ColWidth(24) = 900:            .ColAlignment(24) = flexAlignRightCenter
        .TextArray(25) = "�˻�":        .ColWidth(25) = 900:            .ColAlignment(25) = flexAlignRightCenter
        .TextArray(26) = "���":      .ColWidth(26) = 1000:           .ColAlignment(26) = flexAlignRightCenter
        .TextArray(27) = "���������ڵ�":        .ColWidth(27) = 0
        .TextArray(28) = "������":      .ColWidth(28) = 0
        
        .TextArray(.Cols + 0) = " "
        .TextArray(.Cols + 1) = "�ŷ�ó"
        .TextArray(.Cols + 2) = "ǰ��"
        .TextArray(.Cols + 3) = "������ȣ" & vbCrLf & "��  ��  ��"
        .TextArray(.Cols + 4) = "Order No." & vbCrLf & "��  ��  ��"
        .TextArray(.Cols + 5) = "����"
        .TextArray(.Cols + 6) = "���ַ�"
        .TextArray(.Cols + 7) = "�԰�"
        .TextArray(.Cols + 8) = "�̰�ȹ��"
        .TextArray(.Cols + 9) = "��ȹ��"
        .TextArray(.Cols + 10) = "��ⷮ"
        .TextArray(.Cols + 11) = "�����"
        .TextArray(.Cols + 12) = "������"
        .TextArray(.Cols + 13) = "����"
        .TextArray(.Cols + 14) = "����"
        .TextArray(.Cols + 15) = "S/K"
        .TextArray(.Cols + 16) = "SETT"
        .TextArray(.Cols + 17) = "PEACH"
        .TextArray(.Cols + 18) = "CPB"
        .TextArray(.Cols + 19) = "����"
        .TextArray(.Cols + 20) = "DRY"
        .TextArray(.Cols + 21) = "����"
        .TextArray(.Cols + 22) = "�˻�"
        .TextArray(.Cols + 23) = "����"
        .TextArray(.Cols + 24) = "�հ�"
        .TextArray(.Cols + 25) = "���հ�"
        .TextArray(.Cols + 26) = "���"
        .TextArray(.Cols + 27) = "���������ڵ�"
        .TextArray(.Cols + 28) = "������"

        .ColFormat(6) = "#,##0"
        .ColFormat(7) = "#,##0"
        .ColFormat(8) = "#,##0"
        .ColFormat(9) = "#,##0"
        .ColFormat(10) = "#,##0"
        .ColFormat(11) = "#,##0"
        .ColFormat(12) = "#,##0"
        .ColFormat(13) = "#,##0"
        .ColFormat(14) = "#,##0"
        .ColFormat(15) = "#,##0"
        .ColFormat(16) = "#,##0"
        .ColFormat(17) = "#,##0"
        .ColFormat(18) = "#,##0"
        .ColFormat(19) = "#,##0"
        .ColFormat(20) = "#,##0"
        .ColFormat(21) = "#,##0"
        .ColFormat(22) = "#,##0"
        .ColFormat(23) = "#,##0"
        .ColFormat(24) = "#,##0"
        .ColFormat(25) = "#,##0"
        .ColFormat(26) = "#,##0"
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
 '       .FrozenCols = 5

        For i = 0 To 9
            .MergeCol(i) = True
        Next i
        
        For i = 12 To 23
            .MergeCol(i) = True
        Next i
        .MergeCol(26) = True
        .MergeCol(27) = True
        .MergeCol(28) = True
       
        For i = 1 To .Cols - 1
            .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
            .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
        Next i
        
        For i = 13 To 23
            .ColHidden(i) = True
        Next i
        
        .ColHidden(14) = True
        .ColHidden(15) = True
        .ColHidden(17) = True
        .ColHidden(20) = True
        
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 0
        .Redraw = flexRDDirect
    End With

    '��ޱ���, ��ȹ����, �ŷ�ó, ǰ��, ������ȣ, ������ȣ, ������, ����
    Call SetVSFlexGrid(grdPlanData)
    With grdPlanData
        .Redraw = flexRDNone

        .Row = 0
        .Cols = 11

        .TextArray(1) = "����":       .ColWidth(1) = 500:              .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "��ȹ":       .ColWidth(2) = 500:              .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "���":       .ColWidth(3) = 500:              .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "ǰ��":       .ColWidth(4) = 2600:             .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "������ȣ":   .ColWidth(5) = 1300:             .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "OrderID":    .ColWidth(6) = 1300:             .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "�����":     .ColWidth(7) = 2400:             .ColAlignment(7) = flexAlignLeftCenter
        .TextArray(8) = "����":       .ColWidth(8) = 1000:             .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "����":       .ColWidth(9) = 2000:             .ColAlignment(9) = flexAlignLeftCenter
        .TextArray(10) = "Orderseq":  .ColWidth(10) = 0:               .ColAlignment(10) = flexAlignLeftCenter
        
        .ColHidden(10) = True
        

''         For i = 0 To .Cols - 1
''             nWidth = nWidth + .ColWidth(i)
''         Next i
''        .Width = nWidth

        .ColDataType(1) = flexDTBoolean

        .Redraw = flexRDDirect
    End With
    
    With grdTotal
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByColumn
        .FixedRows = 0
        .Rows = 1
        .Cols = 2
        .ExtendLastCol = True
        
        .RowHeight(0) = 300
        .TextArray(0) = "�հ�":           .ColWidth(0) = 2000:   .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "YD:              .ColWidth(1) = 700:    .ColAlignment(1) = flexAlignRightCenter"
        
        .RowHeight(0) = 300
        .Redraw = flexRDDirect
    End With
End Sub


''Private Sub CheckCount()
''    With grdOrder
''        If .Cell(flexcpChecked, .Row, 1) = flexUnchecked Then
''            .Cell(flexcpChecked, .Row, 1) = flexChecked
''            m_nSelected = m_nSelected + 1
''        Else
''            .Cell(flexcpChecked, .Row, 1) = flexUnchecked
''            m_nSelected = m_nSelected - 1
''        End If
''    End With
''
''    cmdClose.Enabled = IIf(m_nSelected > 0, True, False)
''End Sub


Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub





Private Sub grdPlanData_Click()
    Dim Checked As Boolean
    
    With grdPlanData
        If .Row < .FixedRows Then Exit Sub
        
        If .Col = 1 Then
            Checked = IIf(.Cell(flexcpChecked, .Row, .Col) = flexChecked, False, True)  'üũ�Ǹ� true, üũ������ false
            .Cell(flexcpChecked, .Row, .Col) = Checked
        End If
        optOrder(1).Value = True
        Call ShowData(MakeOrderID(.TextMatrix(.Row, 5), OM_REDUCE), .TextMatrix(.Row, 6), .TextMatrix(.Row, 10))
    End With
    
End Sub

'/********************************************************
' * Description : CPB / Rapid ���� ���԰�ȹ
' * ��       �� : pl_mast ��ȹ���� ��������  select
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    �� ¥        �ۼ���    ����                   �������
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     ������    1.0     �ۼ�
' ********************************************************/
' ��ޱ���, ��ȹ����, �ŷ�ó , ǰ��, ������ȣ, ������ȣ, ������, ����
Private Sub ShowData(ByVal OrderID As String, ByVal OrderNo As String, ByVal OrderSeq As Integer)

    Dim dSql_str As String
    Dim dRS As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    dSql_str = " SELECT EmerClss " & vbCr & _
               "      , PlanClss " & vbCr & _
               "      , Remark, Qty = ISNULL(Qty,0) " & vbCr & _
               "      , PersonName = ISNULL( ( SELECT [Name] " & vbCr & _
               "                                 From mt_person " & _
               "                                WHERE PersonID = AA.PersonID), '' ) " & vbCr & _
               "      , PersonID " & vbCr & _
               "      , Color =  ISNULL( ( SELECT Color " & vbCr & _
               "                             From [OrderColor] DD " & vbCr & _
               "                            Where DD.OrderID = aa.OrderID " & vbCr & _
               "                              AND DD.OrderSeq = AA.OrderSeq), '' ) " & vbCr & _
               "   FROM [pl_mast] AA, [mt_Process] BB " & vbCr & _
               "  WHERE AA.PlanDate = '" & MakeDate(DF_SHORT, dtpPlanDate) & "' " & vbCr & _
               "    AND AA.ProcessID = '" & m_ProcessID & "' " & vbCr & _
               "    AND AA.OrderID = '" & Trim$(OrderID) & "' " & _
               "    AND AA.ProcessID = BB.ProcessID " & _
               "    AND AA.OrderSeq  = " & val(OrderSeq)
            
                   
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount = 1 Then
        txtOrderID = OrderID
        cboEmerClss.ListIndex = Trim(FindItem(cboEmerClss, dRS!EmerClss))
        cboPlanClss.ListIndex = Trim(FindItem(cboPlanClss, dRS!PlanClss))
        txtRemark.Text = Trim(dRS!Remark)
        txtPersonID.Text = dRS!PersonName
        txtPersonID.Tag = dRS!PersonID
        txtColorName.Text = dRS!Color
        txtColorName.Tag = OrderSeq
        txtQty = dRS!Qty
        
        Call SetOrderID(OrderID, OrderNo)
        Call FillGridOrder
    End If
               
    dRS.Close
    Set dRS = Nothing
    
    Exit Sub
    
ErrHandler:
    If Not dRS Is Nothing Then
        Set dRS = Nothing
    End If
    
    Call ErrorBox(Err.Number, "frmPlanCPB.FillPlanCPB", Err.Description)
    
End Sub

Sub SetOrderID(ByVal OrderID As String, ByVal OrderNo As String)
    ' Order No
    If optOrder(0).Value = True Then
        txtOrderID.Text = OrderNo
        txtOrderID.Tag = OrderID
    Else
        txtOrderID.Text = OrderID
        txtOrderID.Tag = OrderNo
    End If
End Sub

'/********************************************************
' * Description : CPB / Rapid ���� ���԰�ȹ
' * ��       �� : pl_mast ��ȹ���� ��������  select
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    �� ¥        �ۼ���    ����                   �������
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     ������    1.0     �ۼ�
' ********************************************************/
' ��ޱ���, ��ȹ����, �ŷ�ó , ǰ��, ������ȣ, ������ȣ, ������, ����
Private Sub FillGrdPlanCPB()
    Dim oPlanCPB As New PlusLib2.CPlanCPB
    Dim rs As Recordset, iProcID$
    Dim nNowRow%, nRowCount%, i%, nTotQty As Long
    
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrHandler
    
    m_bLoading = True

    Set oPlanCPB = New PlusLib2.CPlanCPB
    oPlanCPB.Connection = g_adoCon

    Set rs = oPlanCPB.GetPlanCPBList(MakeDate(DF_SHORT, dtpPlanDate), IIf(cboProcessID.ListIndex = 0, "4000", "4300"))
    
    Set oPlanCPB = Nothing
    
'    m_bSkipEvent = True
    nTotQty = 0
    With grdPlanData
        .Redraw = flexRDNone

        nNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        nRowCount = rs.RecordCount
        For i = 1 To nRowCount
            '-- ������� Progress Barǥ��
            
            '-- ������ grid�� display
            .AddItem CStr(i) & vbTab & vbTab & rs!PlanClss & vbTab & rs!EmerClss & vbTab & _
                rs!ArticleName & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                Trim(rs!OrderNo) & vbTab & Trim(rs!Color) & vbTab & _
                SetCurrency(rs!Qty, 0) & vbTab & rs!Remark & vbTab & rs!OrderSeq
                
                nTotQty = nTotQty + CheckNum(rs!Qty)
                
            '-- 2�ٸ��� Į�� �ֱ�
            If (i Mod 2) = 0 Then
                .Row = i
                .Col = .FixedCols
                .ColSel = .Cols - 1
                .CellBackColor = COLOR_GRIDROW
            End If

            rs.MoveNext
        Next i
        rs.Close
        Set rs = Nothing
        grdTotal.TextMatrix(0, 1) = Format(nTotQty, "##,###,##0 YD")
        
        .Redraw = flexRDDirect
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            
            If .Rows <= nNowRow Then
                .Row = .Rows - 1
            Else
                .Row = nNowRow
            End If
            
            .Col = .FixedCols
            .ColSel = .Cols - 1
            Call ShowData(MakeOrderID(.TextMatrix(.Row, 5), OM_REDUCE), .TextMatrix(.Row, 6), .TextMatrix(.Row, 10))
            
        Else
            .HighLight = flexHighlightNever
         '   grdOrder.Rows = grdOrder.FixedRows
         '   Call ClearData
        End If

'        Call ChangeScroll(0)
        
    End With
    
    
    m_nSelected = 0
'    m_bSkipEvent = False
    
    
    m_bLoading = False
    Screen.MousePointer = vbArrow
    
    If grdPlanData.Rows = grdPlanData.FixedRows Then
        grdOrder.Rows = grdOrder.FixedRows
        grdOrder.HighLight = flexHighlightNever
        Exit Sub
    Else
        Call ShowData(MakeOrderID(grdPlanData.TextMatrix(grdPlanData.Row, 5), OM_REDUCE), grdPlanData.TextMatrix(grdPlanData.Row, 6), grdPlanData.TextMatrix(grdPlanData.Row, 10))
    End If
    
    Exit Sub
ErrHandler:
    m_bLoading = False
    
    Set rs = Nothing
    Set oPlanCPB = Nothing
    
    Screen.MousePointer = vbArrow
    
    Call ErrorBox(Err.Number, "frmPlanCPB.FillGrdPlanCPB", Err.Description)
End Sub


Private Sub grdPlanData_RowColChange()
'    Dim Checked As Boolean
'
'    With grdPlanData
'        If m_bLoading Then Exit Sub
'
'        If .Row < .FixedRows Then Exit Sub
'
'        If .Col = 1 Then
'            Checked = IIf(.Cell(flexcpChecked, .Row, .Col) = flexChecked, False, True)  'üũ�Ǹ� true, üũ������ false
'            .Cell(flexcpChecked, .Row, .Col) = Checked
'        End If
'        optOrder(1).Value = True
'
'        Call ShowData(MakeOrderID(.TextMatrix(.Row, 5), OM_REDUCE), .TextMatrix(.Row, 6), .TextMatrix(.Row, 10))
'    End With
End Sub

Private Sub optOrder_Click(Index As Integer)
    Dim mString As String
    
    chkSearch(0).Caption = optOrder(Index).Caption
    mString = txtOrderID.Text
    
    Select Case Index
    Case 0: txtOrderID.Text = txtOrderID.Tag
    
    Case 1: txtOrderID.Text = txtOrderID.Tag
    End Select
    txtOrderID.Tag = mString
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txtOrderID_Change()
    If Len(txtOrderID) = 0 Then
        txtOrderID.Tag = ""
    End If
End Sub

Private Sub txtOrderID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call FillGridOrder
        Call ClearData
    End If
    
End Sub

