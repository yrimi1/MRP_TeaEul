VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmCodeCode 
   BorderStyle     =   1  '���� ����
   Caption         =   "������ �����ڵ� ����"
   ClientHeight    =   8085
   ClientLeft      =   720
   ClientTop       =   2220
   ClientWidth     =   13395
   Icon            =   "frmCodeCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   13395
   Begin VB.ListBox lstHoldClss 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   30
      TabIndex        =   29
      Top             =   90
      Width           =   1365
   End
   Begin Threed.SSPanel pnlEdit 
      Height          =   4545
      Index           =   0
      Left            =   8400
      TabIndex        =   0
      Top             =   1410
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   8017
      _Version        =   196609
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ �Һз� ]"
      Alignment       =   6
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   4215
         Index           =   2
         Left            =   30
         TabIndex        =   14
         Top             =   300
         Width           =   2670
         _cx             =   4710
         _cy             =   7435
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
   Begin Threed.SSPanel pnlEdit 
      Height          =   4545
      Index           =   1
      Left            =   2820
      TabIndex        =   1
      Top             =   1410
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   8017
      _Version        =   196609
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ ��з� ]"
      Alignment       =   6
      FloodColor      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   4215
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   300
         Width           =   2700
         _cx             =   4762
         _cy             =   7435
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
   Begin Threed.SSPanel pnlBoard 
      Height          =   915
      Left            =   7005
      TabIndex        =   2
      Top             =   30
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1614
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&S)"
         Height          =   780
         Index           =   3
         Left            =   60
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   7
         ToolTipText     =   "�ڷ� ����"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "�߰�(&A)"
         Height          =   780
         Index           =   0
         Left            =   1650
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   6
         ToolTipText     =   "�ڷ� �߰�"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&D)"
         Height          =   780
         Index           =   2
         Left            =   3240
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   5
         ToolTipText     =   "�ڷ� ����"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&U)"
         Height          =   780
         Index           =   1
         Left            =   2445
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   4
         ToolTipText     =   "�ڷ� ����"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "���(&C)"
         Height          =   780
         Index           =   4
         Left            =   855
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   3
         ToolTipText     =   "�ڷ� ���"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin Threed.SSPanel pnlWorkArea 
      Height          =   1350
      Left            =   1470
      TabIndex        =   8
      Top             =   30
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   2381
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txtReason 
         Height          =   345
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   870
         Width           =   4305
      End
      Begin VB.TextBox txtAbReason 
         Height          =   345
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   480
         Width           =   4305
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  '��� ����
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   4980
         MaxLength       =   1
         TabIndex        =   22
         Text            =   "3"
         Top             =   60
         Width           =   375
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  '��� ����
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   4605
         MaxLength       =   1
         TabIndex        =   21
         Text            =   "3"
         Top             =   60
         Width           =   375
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  '��� ����
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   4230
         MaxLength       =   1
         TabIndex        =   20
         Text            =   "3"
         Top             =   60
         Width           =   375
      End
      Begin VB.OptionButton optClss 
         Caption         =   "�Һз�"
         Height          =   180
         Index           =   2
         Left            =   3060
         TabIndex        =   19
         Top             =   150
         Width           =   855
      End
      Begin VB.OptionButton optClss 
         Caption         =   "�ߺз�"
         Height          =   180
         Index           =   1
         Left            =   2085
         TabIndex        =   18
         Top             =   150
         Width           =   855
      End
      Begin VB.OptionButton optClss 
         Caption         =   "��з�"
         Height          =   180
         Index           =   0
         Left            =   1110
         TabIndex        =   17
         Top             =   150
         Width           =   855
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   330
         Index           =   0
         Left            =   60
         TabIndex        =   15
         Top             =   90
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         _Version        =   196609
         Caption         =   "�����ڵ�"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   330
         Index           =   1
         Left            =   60
         TabIndex        =   16
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         _Version        =   196609
         Caption         =   "����  ��"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   330
         Index           =   2
         Left            =   60
         TabIndex        =   24
         Top             =   870
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         _Version        =   196609
         Caption         =   "��������"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Label lblProcID 
         Caption         =   "Label1"
         Height          =   195
         Left            =   2070
         TabIndex        =   27
         Top             =   330
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin Threed.SSPanel pnlEdit 
      Height          =   4545
      Index           =   3
      Left            =   5610
      TabIndex        =   9
      Top             =   1410
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   8017
      _Version        =   196609
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "[ �ߺз� ]"
      Alignment       =   6
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   4215
         Index           =   1
         Left            =   30
         TabIndex        =   13
         Top             =   300
         Width           =   2700
         _cx             =   4762
         _cy             =   7435
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
   Begin Threed.SSPanel pnlEdit 
      Height          =   4545
      Index           =   4
      Left            =   30
      TabIndex        =   11
      Top             =   1410
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   8017
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " �� ���� �ڵ�"
      Alignment       =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VSFlex7LCtl.VSFlexGrid grdProcess 
         Height          =   4215
         Left            =   30
         TabIndex        =   10
         Top             =   300
         Width           =   2700
         _cx             =   4762
         _cy             =   7435
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
   Begin Threed.SSPanel pnlMsg 
      Height          =   420
      Left            =   7020
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   741
      _Version        =   196609
      BackColor       =   65535
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   690
      Left            =   9570
      TabIndex        =   28
      Top             =   5970
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      �ݱ�(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "��  ��/��/�Һз����� ""����Ŭ��"" �Ͻø� �ش� �׸��� ����/������ �� �ֽ��ϴ�."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   6150
      Width           =   7065
   End
End
Attribute VB_Name = "frmCodeCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum GroupIndex
    GI_Large = 0
    GI_Middle = 1
End Enum

Private Const LIMIT_WIDTH1 = 2860
Private Const LIMIT_WIDTH2 = 2410
Private Const LIMIT_WIDTH3 = 3000
Private Const LIMIT_ROW = 18

Dim m_bSkip As Boolean
Dim m_sFlag As String * 1

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmCodeCode = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        
    Me.Move 0, 0, 11270, 7100
 
    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
    cmdExit.MousePointer = ssCustom
    cmdExit.MouseIcon = LoadResPicture("POINTER", vbResCursor)

    Call AddListBox
    pnlCaption(0) = "�ҷ��ڵ�"
    pnlCaption(1) = "�ҷ� ��"
    pnlCaption(2) = "�ҷ�����"
    
    Call ClearData
    Call InitGrid
    Call FillGridProcess
    
    pnlWorkArea.Enabled = False
    m_bSkip = False
End Sub

Private Sub InitGrid()
Dim idx%

    ' �����ڵ� Grid
    Call SetVSFlexGrid(grdProcess)
    With grdProcess
        .ExplorerBar = flexExNone
        .Cols = 3
        
        .TextArray(0) = "":             .ColWidth(0) = 330
        .TextArray(1) = "�����ڵ�":         .ColWidth(1) = 900
        .TextArray(2) = "������":       .ColWidth(2) = LIMIT_WIDTH1
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .Redraw = True
    End With
    
    For idx = 0 To 2
        txtCode(idx) = ""
        
        Call SetVSFlexGrid(grdData(idx))
        With grdData(idx)
            .ExplorerBar = flexExNone
            .Cols = 6:          .Rows = 1
            .FixedCols = 1:     .FixedRows = 1
        
            .TextArray(0) = "":             .ColWidth(0) = 330
            .TextArray(1) = "�ڵ�":         .ColWidth(1) = 450
            .TextArray(2) = "����":         .ColWidth(2) = 0
            .TextArray(3) = "��Ī":         .ColWidth(3) = LIMIT_WIDTH1
            .TextArray(4) = "��ü��Ī":     .ColWidth(4) = 0
            .TextArray(5) = "�ۼ���":       .ColWidth(5) = 0
            
        End With
    Next idx
End Sub

Private Sub AddListBox()
    With lstHoldClss
        .Clear
        .AddItem "�ҷ�����"
        .AddItem "��������"
        .AddItem "Ư�����"
        .AddItem "������"
        .Selected(0) = True
    End With
End Sub

Private Function FillGridProcess() As Boolean
    Dim oCode As PlusLib2.CCode
    Dim rs As ADODB.Recordset
    Dim iLoop%
    Dim sKey As String

    On Error GoTo ErrHandler
    
    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon
        
    Set rs = oCode.GetProcessgroup()
    Set oCode = Nothing
    
    With grdProcess
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & rs!PROCESSID & vbTab & CheckNull(rs!Process)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        grdProcess.Row = grdProcess.FixedRows
    End With
    
    Exit Function

ErrHandler:
    MsgBox CStr(Err.Number) & Err.Description, vbCritical
    Err.Clear
    Set rs = Nothing
    Set oCode = Nothing
End Function

Private Sub grdData_DblClick(Index As Integer)
'    If CInt("0" & m_sFlag) = ID_UPDATE Then
        With grdData(Index)
            If .Row >= .FixedRows And Trim(.TextMatrix(.Row, 0)) <> "" Then
                txtCode(0) = Left(.TextMatrix(.Row, 1), 1)
                txtCode(1) = Mid(.TextMatrix(.Row, 1), 2, 1)
                txtCode(2) = Right(.TextMatrix(.Row, 1), 1)
                txtAbReason = .TextMatrix(.Row, 3)
                txtReason = .TextMatrix(.Row, 4)
                optClss(Index).Value = True
            End If
        End With
'    End If
End Sub

Private Sub grdProcess_Click()
    Dim oCode As PlusLib2.CCode
    Dim rs As ADODB.Recordset
    Dim iLoop%, idx%
    Dim sSizeClss$

    For idx = 0 To 2
        grdData(idx).Rows = grdData(idx).FixedRows
    Next idx
    With grdProcess
        If .Row >= .FixedRows Then
            lblProcID = .TextMatrix(.Row, 1)
        
            Set oCode = New PlusLib2.CCode
            oCode.Connection = g_adoCon
                
            Set rs = oCode.GetHoldList(.TextMatrix(.Row, 1), Left(lstHoldClss.Text, 2), "��", "0", "0", "0")
            Set oCode = Nothing
            
            
            
            Do Until rs.EOF
                grdData(0).AddItem CStr(grdData(0).Rows) & vbTab & rs!holdid & vbTab & rs!holdclss & vbTab & _
                                    rs!abreason & vbTab & rs!reason & vbTab & rs!PersonID
                rs.MoveNext
            Loop
            If rs.RecordCount > 0 Then
                grdData(0).Row = grdData(0).FixedRows
            Else
                txtReason = ""
                txtAbReason = ""
                For idx = 0 To 2
                    txtCode(idx) = ""
                    optClss(idx).Value = False
                Next idx
            End If
            Set rs = Nothing
            Set oCode = Nothing
        End If
    End With
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim bResult As Boolean
        
    On Error GoTo ErrHandler
    '---------------------------------------------------------------------------
    Select Case Index   '[1] �߰�
        Case ID_ADDNEW
            m_sFlag = ID_ADDNEW
            Call AbleTxtCode(True)
            Call ClearData
            Call ChangeMode(Me, False)
            pnlMsg.Caption = LoadResString(302)
            pnlWorkArea.Enabled = True
            
    '---------------------------------------------------------------------------
        Case ID_UPDATE '[2] ����
            m_sFlag = ID_UPDATE
            Call AbleTxtCode(False)
            Call ChangeMode(Me, False)
            pnlMsg.Caption = LoadResString(303)
            pnlWorkArea.Enabled = True
            
    '---------------------------------------------------------------------------
        Case ID_DELETE '[3] ����
            Dim sMsg$
        
            If Trim(txtCode(0)) = "" Or Trim(txtCode(1)) = "" Or Trim(txtCode(2)) = "" Then Exit Sub
'            Call AbleTxtCode(False)
            If optClss(0).Value = True Then
                sMsg = "��з������� ������ �Ͻø� ��/�Һз� ���뵵 �����˴ϴ�"
            ElseIf optClss(1).Value = True Then
                sMsg = "�ߺз������� ������ �Ͻø� �Һз� ���뵵 �����˴ϴ�"
            End If
    
            If MsgBox(sMsg & vbCrLf & vbCrLf & "������ �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "����Ȯ��") = vbYes Then
                m_sFlag = ID_DELETE
                
                If SaveData() Then
                    Call grdProcess_RowColChange
                    Call ChangeMode(Me, True)
                End If
            End If
    '---------------------------------------------------------------------------
        Case ID_SAVE  '[4] ����
            If CheckData() = False Then Exit Sub
            If SaveData() Then
                Call grdProcess_RowColChange
                Call ChangeMode(Me, True)
                Call AbleTxtCode(True)
                m_sFlag = ""
                pnlWorkArea.Enabled = False
            Else
                pnlWorkArea.Enabled = True
            End If

    '---------------------------------------------------------------------------
        Case ID_CANCEL '[5] ���
            m_sFlag = ""
            Call AbleTxtCode(True)
            Call ChangeMode(Me, True)
            pnlWorkArea.Enabled = False
            Call ClearData
            Call grdProcess_RowColChange
            
    End Select
    
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmCodeCode.cmdOperate_Click", Err.Description)
End Sub

Private Sub AbleTxtCode(bFlag As Boolean)
    txtCode(0).Enabled = bFlag
    txtCode(1).Enabled = bFlag
    txtCode(2).Enabled = bFlag
    
    optClss(0).Enabled = bFlag
    optClss(1).Enabled = bFlag
    optClss(2).Enabled = bFlag
    
End Sub

Private Function CheckData() As Boolean

    CheckData = True
    
    If Trim(txtCode(0)) = "0" Or Trim(txtCode(0)) = "" Or Trim(txtCode(1)) = "" Or Trim(txtCode(2)) = "" Then
        CheckData = False
        MsgBox "�ڵ� �Է��� �ùٸ��� �ʽ��ϴ�", vbCritical, "�Է¿���"
        Exit Function
    End If
    If optClss(0).Value = True Then     ' ��з�
        If txtCode(1).Text <> "0" Or txtCode(2).Text <> "0" Then
            CheckData = False
            MsgBox "�ڵ� �Է��� �ùٸ��� �ʽ��ϴ�", vbCritical, "�Է¿���"
            Exit Function
        End If
    End If
    If optClss(1).Value = True Then     ' �ߺз�
        If txtCode(2).Text <> "0" Or txtCode(1).Text = "0" Then
            CheckData = False
            MsgBox "�ڵ� �Է��� �ùٸ��� �ʽ��ϴ�", vbCritical, "�Է¿���"
            Exit Function
        End If
    End If
    If optClss(2).Value = True Then     ' �Һз�
        If txtCode(0).Text = "0" Or txtCode(1).Text = "0" Then
            CheckData = False
            MsgBox "�ڵ� �Է��� �ùٸ��� �ʽ��ϴ�", vbCritical, "�Է¿���"
            Exit Function
        End If
    End If
    If Trim(txtAbReason) = "" Then
        CheckData = False
        MsgBox "��Ī�� �ο����� �ʾҽ��ϴ�", vbCritical, "�Է¿���"
        Exit Function
    End If
    If Trim(lblProcID) = "" Then
        CheckData = False
        MsgBox "������ ���õ��� �ʾҽ��ϴ�", vbCritical, "�Է¿���"
        Exit Function
    End If
End Function

Private Sub ClearData()
Dim idx%

    txtAbReason = ""
    txtReason = ""
'    lblProcID = ""
    For idx = 0 To 2
        optClss(idx).Value = False
        txtCode(idx) = ""
    Next idx
End Sub

Private Sub grdProcess_RowColChange()
    Call grdProcess_Click
End Sub

Private Sub lstHoldClss_Click()

    With lstHoldClss
'        grdProcess.Row = 0
        Call ClearData
        grdData(0).Rows = grdData(0).FixedRows
        grdData(1).Rows = grdData(1).FixedRows
        grdData(2).Rows = grdData(2).FixedRows
        Select Case .ListIndex
            Case 0:
                pnlCaption(0) = "�ҷ��ڵ�"
                pnlCaption(1) = "�ҷ� ��"
                pnlCaption(2) = "�ҷ�����"
            Case 1:
                pnlCaption(0) = "�����ڵ�"
                pnlCaption(1) = "���� ��"
                pnlCaption(2) = "��������"
            Case 2:
                pnlCaption(0) = "Ư���ڵ�"
                pnlCaption(1) = "Ư�� ��"
                pnlCaption(2) = "Ư�⳻��"
            Case 3:
                pnlCaption(0) = "����ڵ�"
                pnlCaption(1) = "��� ��"
                pnlCaption(2) = "�����"
        End Select
        Call grdProcess_Click
    End With
End Sub

Private Sub optClss_Click(Index As Integer)
    Select Case Index
        Case 0:
                txtCode(1) = "0"
                txtCode(2) = "0"
        Case 1:
                txtCode(2) = "0"
    End Select
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
    Call WholeSelect(txtCode(Index))
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index < 2 Then
        txtCode(Index + 1).SetFocus
    Else
        txtAbReason.SetFocus
    End If
End Sub

Private Function SaveData() As Boolean
Dim oCode As PlusLib2.CCode
Dim rs As Recordset
Dim sProcID$, sHoldClss$, sHoldID$
Dim i%, sSize$
Dim bExist As Boolean
    
    On Error GoTo ErrHandler
    
    sProcID = grdProcess.TextMatrix(grdProcess.Row, 1)
    sHoldClss = Left(lstHoldClss.Text, 2)
    
    For i = 0 To 2
        sHoldID = sHoldID & IIf(Trim(txtCode(i)) = "", "0", txtCode(i))
    Next i
    
    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon
    oCode.UserName = g_sUserName
    
    If m_sFlag = ID_ADDNEW Then
        Set rs = oCode.GetHoldCode(sProcID, sHoldClss, sHoldID)
        If rs.RecordCount > 0 Then
            bExist = True
            MsgBox "�ش� �ڵ尡 �̹� �����մϴ�. Ȯ�ιٶ��ϴ�", vbCritical, "���ϵ� �ڵ�"
        End If
        Set rs = Nothing
        
        If bExist = True Then
            Set oCode = Nothing
            SaveData = False
            Exit Function
        End If
    End If
    
    
    If m_sFlag = ID_ADDNEW Then
        SaveData = oCode.AddHoldCode(sProcID, sHoldClss, sHoldID, Trim(txtReason), Trim(txtAbReason), g_sUserName)
    ElseIf m_sFlag = ID_UPDATE Then
        SaveData = oCode.UpdateHoldCode(sProcID, sHoldClss, sHoldID, Trim(txtReason), Trim(txtAbReason), g_sUserName)
    ElseIf m_sFlag = ID_DELETE Then
        If optClss(0).Value = True Then
            sSize = "��"
        End If
        If optClss(1).Value = True Then
            sSize = "��"
        End If
        If optClss(2).Value = True Then
            sSize = "��"
        End If
        
        SaveData = oCode.DeleteHoldCode(sProcID, sHoldClss, sHoldID, sSize)
    End If

    Set oCode = Nothing
    Exit Function
    
ErrHandler:
    Set oCode = Nothing

    Call ErrorBox(Err.Number, "frmCodeCode.SaveData", Err.Description)
End Function

Private Function FillGridData(idx As Integer) As Boolean
    Dim oCode As PlusLib2.CCode
    Dim rs As ADODB.Recordset
    Dim iLoop%
    Dim sSizeClss$
    Dim sBigCode$, sMidCode$


    Select Case idx
        Case 0:     sSizeClss = "��"
                    sBigCode = Left(grdData(0).TextMatrix(grdData(0).Row, 1), 1)
                    grdData(1).Rows = grdData(1).FixedRows
                    grdData(2).Rows = grdData(2).FixedRows
        Case 1:     sSizeClss = "��"
                    sBigCode = Left(grdData(0).TextMatrix(grdData(0).Row, 1), 1)
                    sMidCode = Mid(grdData(1).TextMatrix(grdData(1).Row, 1), 2, 1)
                    grdData(2).Rows = grdData(2).FixedRows
        Case 2:     sSizeClss = "��"
    End Select
    Set oCode = New PlusLib2.CCode
    oCode.Connection = g_adoCon
                
                
    Set rs = oCode.GetHoldList(grdProcess.TextMatrix(grdProcess.Row, 1), Left(lstHoldClss.Text, 2), _
                    sSizeClss, sBigCode, sMidCode, "0")
    Set oCode = Nothing
            
    Do Until rs.EOF
        grdData(idx + 1).AddItem CStr(grdData(idx + 1).Rows) & vbTab & rs!holdid & vbTab & rs!holdclss & vbTab & _
                            rs!abreason & vbTab & rs!reason & vbTab & rs!PersonID
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        grdData(idx + 1).Row = grdData(idx + 1).FixedRows
    End If
    Set rs = Nothing
    Set oCode = Nothing
End Function
















Private Sub ShowData()
    
'    If grdData(1).Rows = grdData(1).FixedRows Then
'        Call ClearData
'        Exit Sub
'    End If
'
'    With grdData(1)
'        txtCode = .TextMatrix(.Row, 1)
'        txtName(0) = .TextMatrix(.Row, 2)
'        txtName(1) = .TextMatrix(.Row, 3)
'    End With
End Sub








Private Sub grdData_RowColChange(Index As Integer)
    txtCode(0) = Left(grdData(Index).TextMatrix(grdData(Index).Row, 1), 1)
    txtCode(1) = Mid(grdData(Index).TextMatrix(grdData(Index).Row, 1), 2, 1)
    txtCode(2) = Right(grdData(Index).TextMatrix(grdData(Index).Row, 1), 1)
    txtAbReason = grdData(Index).TextMatrix(grdData(Index).Row, 3)
    txtReason = grdData(Index).TextMatrix(grdData(Index).Row, 4)
    optClss(Index).Value = True
    Call FillGridData(Index)
End Sub









Private Sub txtCode_LostFocus(Index As Integer)
    If Index = 2 Then
        If optClss(0).Value = True Then
            txtCode(1) = "0"
            txtCode(2) = "0"
        End If
        If optClss(1).Value = True Then
            txtCode(2) = "0"
        End If
    End If
End Sub
