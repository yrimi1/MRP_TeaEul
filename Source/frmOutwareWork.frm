VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmOutwareWork 
   Caption         =   "��� �۾�"
   ClientHeight    =   9585
   ClientLeft      =   315
   ClientTop       =   525
   ClientWidth     =   12750
   Icon            =   "frmOutwareWork.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   12750
   Begin Threed.SSCommand cmdExit 
      Height          =   675
      Left            =   10170
      TabIndex        =   6
      Top             =   8550
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1191
      _Version        =   196609
      Caption         =   "      �ݱ�(&X)"
      PictureAlignment=   1
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   9405
      Left            =   -30
      TabIndex        =   0
      Top             =   -30
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   16589
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   2
      TabCaption(0)   =   " _"
      TabPicture(0)   =   "frmOutwareWork.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSearch"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraOrder(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraOrder(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "pnlName(16)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pnlName(15)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "grdColorSum"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "grdOrder"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "grdColor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdReceive"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSend"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "pnlProgress"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboCom"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDelay"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fraBoxClss"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "_"
      TabPicture(1)   =   "frmOutwareWork.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdOutColorSum"
      Tab(1).Control(1)=   "pnlName(9)"
      Tab(1).Control(2)=   "cmdSave"
      Tab(1).Control(3)=   "cmdRcv"
      Tab(1).Control(4)=   "grdOut"
      Tab(1).Control(5)=   "grdOutColor"
      Tab(1).Control(6)=   "pnlName(8)"
      Tab(1).Control(7)=   "pnlCaption(4)"
      Tab(1).ControlCount=   8
      Begin Threed.SSFrame fraBoxClss 
         Height          =   705
         Left            =   3990
         TabIndex        =   69
         Top             =   60
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1244
         _Version        =   196609
         Begin VB.OptionButton optBoxClss 
            Caption         =   "BOX ���"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   71
            Top             =   390
            Width           =   1245
         End
         Begin VB.OptionButton optBoxClss 
            Caption         =   "ROLL ���"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   70
            Top             =   90
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   795
         Index           =   4
         Left            =   -69510
         TabIndex        =   24
         Top             =   8520
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   1402
         _Version        =   196609
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkBoOutClss 
            Caption         =   "ó����"
            Height          =   285
            Left            =   1680
            TabIndex        =   73
            Top             =   420
            Width           =   975
         End
         Begin VB.ComboBox cboOutClss 
            Height          =   300
            Left            =   60
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   25
            Top             =   420
            Width           =   1545
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   285
            Index           =   3
            Left            =   60
            TabIndex        =   26
            Top             =   90
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   503
            _Version        =   196609
            Caption         =   "�����"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   285
            Index           =   6
            Left            =   1650
            TabIndex        =   72
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   196609
            Caption         =   "������"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin VB.TextBox txtDelay 
         Alignment       =   1  '������ ����
         Height          =   300
         Left            =   6660
         TabIndex        =   18
         Top             =   435
         Width           =   990
      End
      Begin VB.ComboBox cboCom 
         Height          =   300
         Left            =   6660
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   17
         Top             =   60
         Width           =   1005
      End
      Begin Threed.SSPanel pnlProgress 
         Height          =   765
         Left            =   510
         TabIndex        =   14
         Top             =   4110
         Visible         =   0   'False
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   1349
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComctlLib.ProgressBar PBar 
            Height          =   300
            Left            =   60
            TabIndex        =   15
            Top             =   420
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgress 
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   90
            Width           =   3225
         End
      End
      Begin Threed.SSPanel pnlName 
         Height          =   525
         Index           =   8
         Left            =   -74940
         TabIndex        =   11
         Top             =   60
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "���� ��� ����"
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdOutColor 
         Height          =   7425
         Left            =   -69480
         TabIndex        =   8
         Top             =   630
         Width           =   6345
         _cx             =   11192
         _cy             =   13097
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
      Begin VSFlex7LCtl.VSFlexGrid grdOut 
         Height          =   7815
         Left            =   -74940
         TabIndex        =   7
         Top             =   630
         Width           =   5355
         _cx             =   9446
         _cy             =   13785
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
      Begin Threed.SSCommand cmdSend 
         Height          =   675
         Left            =   8400
         TabIndex        =   1
         Top             =   8580
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1191
         _Version        =   196609
         Caption         =   "        ��������(&O)"
         PictureAlignment=   1
      End
      Begin Threed.SSCommand cmdReceive 
         Height          =   675
         Left            =   10215
         TabIndex        =   2
         Top             =   90
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1191
         _Version        =   196609
         Caption         =   "����ڷ����"
      End
      Begin VSFlex7LCtl.VSFlexGrid grdColor 
         Height          =   5595
         Left            =   3990
         TabIndex        =   3
         Top             =   2490
         Width           =   7875
         _cx             =   13891
         _cy             =   9869
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
      Begin VSFlex7LCtl.VSFlexGrid grdOrder 
         Height          =   5955
         Left            =   60
         TabIndex        =   4
         Top             =   2490
         Width           =   3885
         _cx             =   6853
         _cy             =   10504
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
      Begin VSFlex7LCtl.VSFlexGrid grdColorSum 
         Height          =   330
         Left            =   3990
         TabIndex        =   5
         Top             =   8100
         Width           =   7905
         _cx             =   13944
         _cy             =   582
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
      Begin Threed.SSCommand cmdRcv 
         Height          =   675
         Left            =   -74910
         TabIndex        =   9
         Top             =   8580
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1191
         _Version        =   196609
         Caption         =   "        �ڷ����(&R)"
         PictureAlignment=   1
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   675
         Left            =   -66600
         TabIndex        =   10
         Top             =   8580
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1191
         _Version        =   196609
         Caption         =   "      ����(&S)"
         PictureAlignment=   1
      End
      Begin Threed.SSPanel pnlName 
         Height          =   525
         Index           =   9
         Left            =   -69480
         TabIndex        =   12
         Top             =   60
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   926
         _Version        =   196609
         Caption         =   "Color�� ��� ����"
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdOutColorSum 
         Height          =   330
         Left            =   -69480
         TabIndex        =   13
         Top             =   8100
         Width           =   6345
         _cx             =   11192
         _cy             =   582
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
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   15
         Left            =   5460
         TabIndex        =   19
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "��� ��Ʈ"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   300
         Index           =   16
         Left            =   5460
         TabIndex        =   20
         Top             =   435
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "���� �ð�"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSFrame fraOrder 
         Height          =   705
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   8550
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1244
         _Version        =   196609
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   120
            Width           =   1185
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "���� ��ȣ"
            Height          =   180
            Index           =   1
            Left            =   90
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   405
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin Threed.SSFrame fraOrder 
         Height          =   1605
         Index           =   1
         Left            =   3990
         TabIndex        =   27
         Top             =   810
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   2831
         _Version        =   196609
         Begin VB.TextBox txtName 
            Alignment       =   1  '������ ����
            Height          =   315
            Index           =   7
            Left            =   3870
            TabIndex        =   37
            Text            =   "0"
            Top             =   1230
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Alignment       =   1  '������ ����
            Height          =   315
            Index           =   6
            Left            =   1260
            TabIndex        =   36
            Text            =   "0"
            Top             =   1230
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Alignment       =   1  '������ ����
            Height          =   315
            Index           =   5
            Left            =   6480
            TabIndex        =   35
            Text            =   "0"
            Top             =   450
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Index           =   4
            Left            =   3870
            TabIndex        =   34
            Top             =   450
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Index           =   3
            Left            =   1260
            TabIndex        =   33
            Top             =   450
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Index           =   2
            Left            =   6480
            TabIndex        =   32
            Top             =   90
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Index           =   1
            Left            =   3870
            TabIndex        =   31
            Top             =   90
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Index           =   0
            Left            =   1260
            TabIndex        =   30
            Top             =   90
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Alignment       =   1  '������ ����
            Height          =   315
            Index           =   8
            Left            =   1260
            TabIndex        =   29
            Text            =   "0"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Alignment       =   1  '������ ����
            Height          =   315
            Index           =   9
            Left            =   3870
            TabIndex        =   28
            Text            =   "0"
            Top             =   840
            Width           =   1335
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   38
            Top             =   90
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "Order No."
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   1
            Left            =   2670
            TabIndex        =   39
            Top             =   90
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "�� �� ó"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   2
            Left            =   5280
            TabIndex        =   40
            Top             =   90
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "ǰ ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   41
            Top             =   450
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "�� �� ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   4
            Left            =   2670
            TabIndex        =   42
            Top             =   450
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "�� �� ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   5
            Left            =   5280
            TabIndex        =   43
            Top             =   450
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "���� ����"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   6
            Left            =   60
            TabIndex        =   44
            Top             =   1230
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "��� ����"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   7
            Left            =   2670
            TabIndex        =   45
            Top             =   1230
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "��� ����"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   10
            Left            =   60
            TabIndex        =   46
            Top             =   840
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "�԰� ����"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   11
            Left            =   2670
            TabIndex        =   47
            Top             =   840
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "�԰� ����"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSFrame fraSearch 
         Height          =   2385
         Left            =   60
         TabIndex        =   48
         Top             =   60
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   4207
         _Version        =   196609
         Begin VB.CommandButton cmdTerm 
            Caption         =   "����"
            Height          =   315
            Index           =   0
            Left            =   345
            MousePointer    =   99  '����� ����
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   465
            Width           =   600
         End
         Begin VB.CommandButton cmdTerm 
            Caption         =   "�ݿ�"
            Height          =   315
            Index           =   1
            Left            =   345
            MousePointer    =   99  '����� ����
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   825
            Width           =   600
         End
         Begin VB.TextBox txtSearch 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   1440
            TabIndex        =   52
            Top             =   1620
            Width           =   1905
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "�˻�(&F)"
            Height          =   780
            Left            =   2985
            MousePointer    =   99  '����� ����
            Style           =   1  '�׷���
            TabIndex        =   51
            ToolTipText     =   "�ڷ� ����"
            Top             =   135
            Width           =   780
         End
         Begin VB.TextBox txtSearch 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1440
            TabIndex        =   50
            Top             =   1230
            Width           =   1905
         End
         Begin VB.TextBox txtSearch 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   1440
            TabIndex        =   49
            Top             =   1995
            Width           =   1905
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   1
            Left            =   3390
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1230
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
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   0
            Left            =   1005
            TabIndex        =   56
            Top             =   465
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            Format          =   113639425
            CurrentDate     =   36871
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   1
            Left            =   1005
            TabIndex        =   57
            Top             =   825
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            Format          =   113639425
            CurrentDate     =   36871
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   1230
            Width           =   1260
            _ExtentX        =   2223
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
               Caption         =   "�� �� ó"
               Height          =   180
               Index           =   1
               Left            =   60
               TabIndex        =   59
               Top             =   60
               Width           =   975
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   1605
            Width           =   1260
            _ExtentX        =   2223
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
               TabIndex        =   61
               Top             =   60
               Width           =   1185
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   62
            Top             =   90
            Width           =   1260
            _ExtentX        =   2223
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
               Caption         =   "���� ����"
               Height          =   240
               Index           =   0
               Left            =   45
               TabIndex        =   63
               Top             =   45
               Width           =   1080
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   5
            Left            =   120
            TabIndex        =   64
            Top             =   1980
            Width           =   1260
            _ExtentX        =   2223
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
               Caption         =   "������ȣ"
               Height          =   180
               Index           =   3
               Left            =   60
               TabIndex        =   65
               Top             =   60
               Width           =   1185
            End
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   2
            Left            =   3390
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   1620
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
         Begin VB.Label lblLabel 
            Alignment       =   2  '��� ����
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   2295
            TabIndex        =   67
            Top             =   540
            Width           =   360
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  '��� ����
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   2295
            TabIndex        =   66
            Top             =   915
            Width           =   360
         End
      End
   End
   Begin MSCommLib.MSComm comOut 
      Left            =   12000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      OutBufferSize   =   2048
      RTSEnable       =   -1  'True
   End
End
Attribute VB_Name = "frmOutwareWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIMIT_WIDTH1 = 1180
Private Const LIMIT_WIDTH2 = 1780
Private Const LIMIT_WIDTH3 = 1075
Private Const LIMIT_WIDTH4 = 1540
Private Const LIMIT_ROW1 = 21
Private Const LIMIT_ROW2 = 20
Private Const LIMIT_ROW3 = 28
Private Const LIMIT_ROW4 = 26

Private m_sRcvBuf As String

Private Type TRcvData
    sManagerID As String * 10    ' ������ȣ
    sBoxClss As String * 1       ' Roll or Box (0,1)
    sBoxNo As String * 4         ' Box ��ȣ
    nRollCnt As Integer          ' Roll ��
    aBarCode(0 To 27) As String * 28 ' BarCode ����
    sPackingDate As String * 12
End Type

Private Type TSndData
    OrderID As String * 10
    OrderNo As String * 12
    BoxClss As String * 1
    OrderQty As String * 7
    OrderOut As String * 7
End Type

Private Type TColor
    Color     As String * 12    ' �����
    OrderQty  As String * 6     ' ���ַ�
    OutQty    As String * 6     ' ���
End Type

Private TOrder As TSndData
Private aSendColor() As TColor


Private Sub chkSearch_Click(Index As Integer)
    If Index = 0 Then '[0] �������� ����
        If chkSearch(0).Value = vbChecked Then
            dtpDate(0).Enabled = True
            dtpDate(1).Enabled = True
        Else
            dtpDate(0).Enabled = False
            dtpDate(1).Enabled = False
        End If
    Else '[1, 2] �ŷ�ó, ������ȣ ����
        If chkSearch(Index).Value = vbChecked Then
            txtSearch(Index).Enabled = True
            txtSearch(Index).SetFocus
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = True
            End If
        Else
            txtSearch(Index).Enabled = False
            If Index = 1 Or Index = 2 Then
                cmdFind(Index).Enabled = False
            End If
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    If tabMain.Tab Then
        If cmdSave.Tag = "RCV_TRUE" Then
            If MsgBox("����� �ڷḦ �������� �����̽��ϴ�." & vbCrLf & "�׷��� �����Ͻðڽ��ϱ�?" & vbCrLf & "'��'�� �����ø� ��ĵ�� ������ �������ϴ�.", vbInformation + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        tabMain.Tab = 0
    Else
        Unload Me
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 1 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
    ElseIf Index = 2 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(Index))
    End If
End Sub

Private Sub cmdRcv_Click()
    Dim RcvData As TRcvData, i%, nRowCnt%, nLoopCnt%
    Dim nCnt&, nSeqNo&, sChk_Input As String * 1
    Dim bSucess As Boolean
    
    On Error GoTo ErrHandler
    
    If cmdSave.Tag = "RCV_TRUE" Then
        If MsgBox("�̹� ����� �ڷḦ �������� �����̽��ϴ�." & vbCrLf & "�������� ���� �ڷ�� �սǵ˴ϴ�. ����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "����!!") = vbNo Then
            Exit Sub
        End If
    End If
    cmdSave.Tag = "RCV_FALSE"
    cmdSave.Enabled = False
    bSucess = False
    
    Call InitComm
    
    MsgBox "DT-900(���ܸ���)�� ���� �غ��ϼž� �մϴ�. " & vbCrLf & vbCrLf & _
            "���ܸ��⿡�� ""�ڷ� �ۼ���"" ���������� Ȯ���Ͻð�" & vbCrLf & vbCrLf & _
            "Ȯ�ι�ư�� ��������..", vbInformation + vbOKOnly, "���� �غ�!!"
    
    comOut.Output = STX & "@20   " & ETX ' �ۼ��� �䱸
    comOut.InBufferCount = 0
    nCnt = 0
    Do While nCnt < 300000
        If CheckRcv Then
            Exit Do
        Else
            nCnt = nCnt + 1
        End If
    Loop
    If nCnt >= 300000 Then
        MsgBox "���ܸ���� �������� �۽��� �ȵ˴ϴ�." & vbCrLf & vbCrLf & "�ٽ� �õ��Ͽ� �ֽʽÿ�.", vbCritical + vbOKOnly, "���� ����!!"
        Exit Sub
    End If
      
    Call Sleep(txtDelay)
    sChk_Input = comOut.Input
    If sChk_Input <> ACK Then
        MsgBox "���ܸ���� ����� ������ �߻��Ͽ����ϴ�." & vbCrLf & vbCrLf & "�ٽ� �õ��Ͽ� �ֽʽÿ�.", vbCritical + vbOKOnly, "���� ����!!"
        Exit Sub
    End If
      
    nRowCnt = 0
    ' ����Ÿ ���۹���
            
    With grdOut
        .Redraw = flexRDDirect
       
        .Rows = .FixedRows
    
        For nLoopCnt = 0 To 8000
            If RcvFrame = 1 Then
                If Left(m_sRcvBuf, 1) = "9" Then
                    MsgBox "�����ڷḦ ��� ���� �޾ҽ��ϴ�.", vbCritical + vbOKOnly, "���� �Ϸ�!!"
                    cmdSave.Tag = "RCV_TRUE"
                    cmdSave.Enabled = True
                    bSucess = True
                    Close #6
                    Exit For
                End If
                
                .AddItem CStr(.Rows) & vbTab & Mid(m_sRcvBuf, 7, 10) & vbTab & _
                                   Mid(m_sRcvBuf, 17, 2) & vbTab & CInt(Mid(m_sRcvBuf, 19, 4)) & vbTab & _
                                   CInt(Mid(m_sRcvBuf, 23, 4)) & vbTab & Trim(Mid(m_sRcvBuf, 31, 5)) & vbTab & CInt(Mid(m_sRcvBuf, 36, 3))
                                   
            Else
                MsgBox "����ڷḦ ���� �޴��� ������ �߻��Ǿ����ϴ�." & vbCrLf & "�ٽ� �����Ͽ� �ֽʽÿ�.", vbCritical + vbOKOnly, "���� ����!!"
                bSucess = False
                Exit For
            End If
        Next nLoopCnt
        
        If .Rows = 1 Then
            .HighLight = flexHighlightNever
        Else
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
            .Col = 1
            .ColSel = .Cols - 1
        End If
        Call ChangeScroll(3)
        
        .Redraw = flexRDDirect
    End With
    
    Call EndComm
    
    If bSucess Then
        Call FillGridOutColor
    Else
        grdOut.Rows = grdOut.FixedRows
        grdOutColor.Rows = grdOutColor.FixedRows
    End If
    
    Exit Sub
    
ErrHandler:
    Call EndComm
    Call ErrorBox(Err.Number, "frmOutWareWork.CmdRcv", Err.Description)
End Sub

Private Sub cmdReceive_Click()
    tabMain.Tab = 1
    
    cmdSave.Enabled = False
    cmdSave.Tag = "RCV_FALSE"
    
    grdOut.Rows = grdOut.FixedRows
    grdOutColor.Rows = grdOutColor.FixedRows
    
    With grdOutColorSum
        .TextArray(1) = 0
        .TextArray(2) = 0
    End With
End Sub

Private Sub cmdSave_Click()
    Dim ow       As PlusLib2.TOUTWARE
    Dim owSub()  As PlusLib2.TOUTWARESUB
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    If grdOut.Rows = grdOut.FixedRows Then
        MsgBox "������ �ڷᰡ �����ϴ�.", vbInformation
        Exit Sub
    End If
   
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Set rs = oOutware.GetOrderOne(grdOut.TextMatrix(1, 1))
    Set oOutware = Nothing
   
    With ow
        ow.OrderID = grdOut.TextMatrix(1, 1)
        ow.OutClss = cboOutClss.ItemData(cboOutClss.ListIndex)
        ow.WorkID = rs!WorkID
        ow.ExchRate = rs!ExchRate
        ow.UnitPrice = rs!UnitPrice
        ow.OutCustom = ""
        ow.LossRate = 0
        ow.LossQty = 0
        ow.OutDate = Format(Date, "YYYYMMDD")
        ow.ResultDate = ""
        ow.OutTime = ""
        ow.LoadTime = Format(time, "HHMM")
        ow.BoOutClss = IIf(chkBoOutClss.Value, "*", "")
        ow.OutRoll = 0
        ow.OutQty = 0
    End With

    With grdOut
        ReDim owSub(.Rows - 2)
        
        For i = .FixedRows To .Rows - 1
            owSub(i - 1).OrderID = ow.OrderID
            owSub(i - 1).OutSubSeq = i
            owSub(i - 1).OrderSeq = CInt(.TextMatrix(i, 2))
            owSub(i - 1).BoxNo = .TextMatrix(i, 3)
            owSub(i - 1).RollSeq = .TextMatrix(i, 4)
            owSub(i - 1).LotNo = .TextMatrix(i, 5)
            owSub(i - 1).OutQty = CSng(.TextMatrix(i, 6))

            ow.OutRoll = ow.OutRoll + 1
            ow.OutQty = ow.OutQty + CSng(.TextMatrix(i, 6))
        Next i
    End With
    
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    oOutware.UserName = g_sUserName

    If oOutware.AddNewOutwareHT(ow, owSub) Then
        MsgBox "��� �ڷḦ �����߽��ϴ�.", vbInformation + vbOKOnly, "���� �Ϸ�!!"

        grdOut.Rows = grdOut.FixedRows

        cmdSave.Tag = "RCV_FALSE"
        cmdSave.Enabled = False
    End If
    Set oOutware = Nothing
    
    Exit Sub
    
ErrHandler:
    Set oOutware = Nothing

    Call ErrorBox(Err.Number, "frmOutware.SaveEditData", Err.Description)
End Sub

Private Sub cmdSearch_Click()
    Call FillGridOrder
End Sub

Private Sub cmdSend_Click()
    Dim i%, j%, nCnt&
    Dim sChk_Input As String * 1
    Dim sDatalen As String * 4
    
    If grdOrder.Rows = grdOrder.FixedRows Then
        MsgBox "������ Order�� ���� �����ϼ���..", vbInformation + vbOKOnly
        Exit Sub
    End If
   
    Call InitComm
   
    MsgBox "DT-900(���ܸ���)�� ���� �غ��ϼž� �մϴ�. " & vbCrLf & vbCrLf & _
            "���ܸ��⿡�� ""�ڷ� �ۼ���"" ���������� Ȯ���Ͻð�" & vbCrLf & vbCrLf & _
            "Ȯ�ι�ư�� ��������..", vbInformation + vbOKOnly, "�˸�!!"
    
    Screen.MousePointer = vbHourglass
    
    comOut.InBufferCount = 0
    comOut.Output = STX & "@10   " & ETX ' �ۼ��� �䱸

    nCnt = 0
    Do While nCnt < 300000
        If CheckRcv Then
            Exit Do
        Else
            nCnt = nCnt + 1
        End If
    Loop
    If nCnt >= 300000 Then
        Screen.MousePointer = vbArrow
        MsgBox "���ܸ���� �������� ������ �ȵ˴ϴ�..." & vbCrLf & vbCrLf & "�ٽ� �õ��Ͻʽÿ�..", vbCritical + vbOKOnly
        
        Exit Sub
    End If
      
    Call Sleep(CInt(txtDelay))
    sChk_Input = comOut.Input
    If sChk_Input <> ACK Then
        Screen.MousePointer = vbArrow
        MsgBox "��� �ܸ���� ����� ������ �߻��Ͽ����ϴ�." & vbCrLf & vbCrLf & "�ٽ� �õ��Ͻʽÿ�..", vbCritical + vbOKOnly
        Exit Sub
    End If
      
    TOrder.OrderID = MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 1), OM_REDUCE)  ' ������ȣ
    TOrder.OrderNo = grdOrder.TextMatrix(grdOrder.Row, 2) '������ȣ
    TOrder.BoxClss = IIf(optBoxClss(0).Value = True, "0", "1") '��� ���� 0 : Roll 1 : Box
    TOrder.OrderQty = grdOrder.TextMatrix(grdOrder.Row, 8)        '���� ����
    TOrder.OrderOut = grdOrder.TextMatrix(grdOrder.Row, 4)        '�� ��� ����
    
    sDatalen = str((UBound(aSendColor) + 1) * 24 + 37)
    comOut.Output = STX & "81" & sDatalen & TOrder.OrderID & TOrder.OrderNo & TOrder.BoxClss & Format(TOrder.OrderQty, "000000#") & Format(TOrder.OrderOut, "000000#")
    
    pnlProgress.Visible = True
    For i = 0 To UBound(aSendColor)
        comOut.Output = StrConv(MidB(StrConv(aSendColor(i).Color, vbFromUnicode), 1, 16), vbUnicode) ' �����
        comOut.Output = aSendColor(i).OrderQty  ' ���ַ�
        comOut.Output = aSendColor(i).OutQty ' ���
        
        
        lblProgress.Caption = "���� ������ : " & CInt(IIf((i + 1) * 100 / (grdColor.Rows - 1) > 100, 100, (i + 1) * 100 / (grdColor.Rows - 1))) & "%"
        PBar.Value = IIf((i + 1) * 100 / (grdColor.Rows - 1) > 100, 100, (i + 1) * 100 / (grdColor.Rows - 1))
        
        Call Sleep(CInt(txtDelay))

        DoEvents
    Next i
    comOut.Output = ETX

    nCnt = 0
    Do While nCnt < 300000
        If CheckRcv Then
            Exit Do
        Else
            nCnt = nCnt + 1
        End If
    Loop
    Screen.MousePointer = vbArrow
    
    pnlProgress.Visible = False
    If nCnt >= 600000 Then
        MsgBox "Order ����Ÿ ������ ������ �߻��Ͽ����ϴ�." & vbCrLf & vbCrLf & "�ٽ� �����Ͻʽÿ�..", vbCritical + vbOKOnly
    Else
        sChk_Input = comOut.Input
        MsgBox "Order ����Ÿ�� ��� �ܸ��⿡ �����Ͽ����ϴ�.", vbInformation + vbOKOnly
    End If
    
    Call EndComm
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[1] ����
        dtpDate(0) = Date
        dtpDate(1) = Date
    ElseIf Index = 1 Then '[2] �ݿ�
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = Date
    End If
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call NextFocus
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11975, 9660
    
    Call InitGrid
    
    dtpDate(0) = Date
    dtpDate(1) = Date
    
    For i = 1 To 2
        cmdFind(i).Enabled = False
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
    Next i
    
    cmdRcv.Picture = LoadResPicture("COMM", vbResIcon)
    cmdSend.Picture = LoadResPicture("COMM", vbResIcon)
    cmdExit.Picture = LoadResPicture("QUIT", vbResIcon)
    cmdSave.Picture = LoadResPicture("SAVE", vbResIcon)
    cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)
    
    With cboOutClss
        .AddItem "1. �������"
        .ItemData(0) = 1
        .AddItem "2. �������"
        .ItemData(1) = 2
        .AddItem "3. ������"
        .ItemData(2) = 3
        .AddItem "3. �����ҷ�"
        .ItemData(3) = 4
        .AddItem "3. �����ҷ�"
        .ItemData(4) = 5
        .AddItem "3. SAMPLE,�ð���"
        .ItemData(5) = 6
        .AddItem "3. �������"
        .ItemData(6) = 7

        .ListIndex = 0
    End With
    
    With cboCom
        .AddItem "COM 1"
        .AddItem "COM 2"
    End With
    
    cboCom.ListIndex = CInt(GetSetting(LoadResString(100), Me.Name, "ComPort", "1")) - 1
    txtDelay = GetSetting(LoadResString(100), Me.Name, "Delay", "600")
    chkSearch(0).Value = vbChecked
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSetting(LoadResString(100), Me.Name, "ComPort", cboCom.ListIndex + 1)
    Call SaveSetting(LoadResString(100), Me.Name, "Delay", txtDelay)
End Sub

Private Sub grdOrder_RowColChange()
    Call ShowData
End Sub

Private Sub optOrder_Click(Index As Integer)
    With grdOrder
        If optOrder(0).Value Then
            .ColWidth(1) = 0
            .ColWidth(2) = 1350
            chkSearch(2).Caption = "Order No."
            pnlName(0).Caption = "Order No"
        Else
            .ColWidth(1) = 1350
            .ColWidth(2) = 0
            chkSearch(2).Caption = "������ȣ"
            pnlName(0).Caption = "������ȣ"
        End If
    End With
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Index = 1 Then
        Call ReturnCode(LG_CUSTOM, , False, txtSearch(1))
    ElseIf KeyAscii = vbKeyReturn And Index = 2 Then
        Call NextFocus
    End If
End Sub

Private Sub InitGrid()
    Call SetVSFlexGrid(grdOrder)
    With grdOrder
        .Redraw = flexRDNone
        .Cols = 13
            
        .TextArray(0) = "����":         .ColWidth(0) = 405
        .TextArray(1) = "������ȣ":     .ColWidth(1) = 1350:            .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "Order No":     .ColWidth(2) = 0:               .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "�ŷ�ó":       .ColWidth(3) = LIMIT_WIDTH1:    .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "������":       .ColWidth(4) = 870:             .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "�������":     .ColWidth(5) = 0
        .TextArray(6) = "ǰ��":         .ColWidth(6) = 0
        .TextArray(7) = "������":       .ColWidth(7) = 0
        .TextArray(8) = "���ַ�":       .ColWidth(8) = 0
        .TextArray(9) = "������":       .ColWidth(9) = 0
        .TextArray(10) = "UNIT":        .ColWidth(10) = 0
        .TextArray(11) = "�԰�����:     .colwidth(11) = 0"
        .TextArray(12) = "�԰����:     .colwidth(12) = 0"

        .ColFormat(4) = "#,##0"
        
        .Redraw = flexRDDirect
    End With
    
    Call SetVSFlexGrid(grdColor)
    With grdColor
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 7
        
        .TextArray(0) = "����":         .ColWidth(0) = 505:             .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "�����":       .ColWidth(1) = LIMIT_WIDTH2:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "Design No.":   .ColWidth(2) = 1500:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "���� ����":    .ColWidth(3) = 1000:            .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "��� ����":    .ColWidth(4) = 1000:            .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "��� ����":    .ColWidth(5) = 1000:            .ColAlignment(5) = flexAlignRightCenter
        .TextArray(6) = "��� �ܷ�":    .ColWidth(6) = 1000:            .ColAlignment(6) = flexAlignRightCenter
        
        .TextArray(0) = .TextArray(0) & vbCrLf & "����"
        
        .ColFormat(3) = "#,##0"
        .ColFormat(4) = "#,##0"
        .ColFormat(5) = "#,##0"
        .ColFormat(6) = "#,##0"
        
        .Redraw = flexRDDirect
    End With

    With grdColorSum
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightNever
'        .SelectionMode = flexSelectionByColumn
        .FixedRows = 0
        .Rows = 1
        .Cols = 5
        .ExtendLastCol = True
        
        .RowHeight(0) = 275
        
        .TextArray(0) = "�հ�":         .ColWidth(0) = 3785:        .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "0":            .ColWidth(1) = 1000:        .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "0":            .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignRightCenter
        .TextArray(3) = "0":            .ColWidth(3) = 1000:        .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "0":            .ColWidth(4) = 1000:        .ColAlignment(4) = flexAlignRightCenter

        .ColFormat(1) = "#,##0"
        .ColFormat(2) = "#,##0"
        .ColFormat(3) = "#,##0"
        .ColFormat(4) = "#,##0"
        
        .Redraw = flexRDDirect
    End With

    Call SetVSFlexGrid(grdOut)
    With grdOut
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 7
        '5265  500
        .TextArray(0) = ""
        .TextArray(1) = "������ȣ":         .ColWidth(1) = 1100:         .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "����" & vbCrLf & "����": .ColWidth(2) = 600:    .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "Box NO":           .ColWidth(3) = 800:          .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "����ȣ":           .ColWidth(4) = 800:          .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "Lot No":           .ColWidth(5) = 800:         .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "����":             .ColWidth(6) = LIMIT_WIDTH3:   .ColAlignment(6) = flexAlignCenterCenter
        
        .Redraw = flexRDDirect
    End With
    
    Call SetVSFlexGrid(grdOutColor)
    With grdOutColor
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 6
        
        .TextArray(0) = "����":         .ColWidth(0) = 505:             .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "�����":       .ColWidth(1) = LIMIT_WIDTH4:    .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "Design No.":   .ColWidth(2) = 1300:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "���� ����":    .ColWidth(3) = 1000:            .ColAlignment(3) = flexAlignRightCenter
        .TextArray(4) = "��� ����":    .ColWidth(4) = 1000:            .ColAlignment(4) = flexAlignRightCenter
        .TextArray(5) = "���� ����":    .ColWidth(5) = 1000:            .ColAlignment(5) = flexAlignRightCenter
        
        .TextArray(0) = .TextArray(0) & vbCrLf & "����"
        
        .ColFormat(3) = "#,##0"
        .ColFormat(4) = "#,##0"
        .ColFormat(5) = "#,##0"
        
        .Redraw = flexRDDirect
    End With

    With grdOutColorSum
        .Redraw = flexRDNone
        
        .HighLight = flexHighlightNever
'        .SelectionMode = flexSelectionByColumn
        .FixedRows = 0
        .Rows = 1
        .Cols = 3
        .ExtendLastCol = True
        
        .RowHeight(0) = 275
        
        .TextArray(0) = "�հ�":         .ColWidth(0) = 4345:        .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "0":            .ColWidth(1) = 1000:        .ColAlignment(1) = flexAlignRightCenter
        .TextArray(2) = "0":            .ColWidth(2) = 1000:        .ColAlignment(2) = flexAlignRightCenter

        .ColFormat(1) = "#,##0"
        .ColFormat(2) = "#,##0"
        
        .Redraw = flexRDDirect
    End With
    
    
End Sub

Private Sub FillGridOrder()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    oOutware.UserName = g_sUserName
    
    Set rs = oOutware.GetOrderList(IIf(chkSearch(0).Value = vbChecked, 1, 0), _
                MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                IIf(chkSearch(1).Value = vbChecked, 1, 0), txtSearch(1).Tag, _
                IIf(chkSearch(2).Value = vbChecked, 1, 0), txtSearch(2).Tag, _
                IIf(chkSearch(3).Value = vbChecked, IIf(optOrder(0).Value, 2, 1), 0), txtSearch(3))
    Set oOutware = Nothing
        
    With grdOrder
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                rs!kCustom & vbTab & rs!OutQty & vbTab & rs!OutRoll & vbTab & rs!Article & vbTab & _
                CheckNull(rs!WorkName) & vbTab & rs!OrderQty & vbTab & rs!WorkWidth & vbTab & rs!UnitClss & vbTab & _
                rs!StuffInRoll & vbTab & rs!StuffInQty

            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    
        Call ChangeScroll(0)
        
        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways
            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
                        
            Call ShowData
        Else
            .HighLight = flexHighlightNever
                    
            Call ClearData
            MsgBox LoadResString(203), vbInformation
        End If
        .Redraw = flexRDDirect
    End With
        
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oOutware = Nothing
    
    Call ErrorBox(Err.Number, "frmOutWareWork.FillGridOrder", Err.Description)
End Sub

Private Sub FillGridColor()
    Dim oOutware As PlusLib2.COutWare
    Dim rs     As ADODB.Recordset
    Dim i%, nOrderQtySum&, nOutRollSum&, nOutQtySum&, nOutLeftSum&
    
    On Error GoTo ErrHandler

    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Set rs = oOutware.GetOrderSubTotal(MakeOrderID(grdOrder.TextMatrix(grdOrder.Row, 1), OM_REDUCE))
    Set oOutware = Nothing

    nOrderQtySum = 0
    nOutRollSum = 0
    nOutQtySum = 0
    With grdColor
        .Redraw = flexRDNone

        .Rows = .FixedRows
        
        ReDim aSendColor(rs.RecordCount - 1)
        For i = 0 To rs.RecordCount - 1
            .AddItem rs!OrderSeq & vbTab & rs!Color & vbTab & CheckNull(rs!DesignNO) & vbTab & rs!ColorQty & vbTab & _
                    rs!OutRoll & vbTab & rs!OutQty & vbTab & rs!ColorQty - rs!OutQty

            aSendColor(i).Color = rs!Color
            aSendColor(i).OrderQty = Format(rs!ColorQty, "00000#")
            aSendColor(i).OutQty = Format(rs!OutQty, "00000#")

            nOrderQtySum = nOrderQtySum + rs!ColorQty
            nOutRollSum = nOutRollSum + rs!OutRoll
            nOutQtySum = nOutQtySum + rs!OutQty
            nOutLeftSum = nOutLeftSum + rs!ColorQty - rs!OutQty
            
            rs.MoveNext
        Next i

        rs.Close
        Set rs = Nothing

        If .Rows > .FixedRows Then
            .HighLight = flexHighlightAlways

            .Row = .FixedRows
            .Col = .FixedCols
            .ColSel = .Cols - 1
        End If

        Call ChangeScroll(1)

        .Redraw = flexRDDirect
    End With

    With grdColorSum
        .TextArray(1) = nOrderQtySum
        .TextArray(2) = nOutRollSum
        .TextArray(3) = nOutQtySum
        .TextArray(4) = nOutLeftSum
    End With
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oOutware = Nothing
    Call ErrorBox(Err.Number, "frmOutware.FillGridColor", Err.Description)
End Sub


Private Sub FillGridOutColor()
    Dim oOutware As PlusLib2.COutWare
    Dim rs As ADODB.Recordset
    Dim nTotal&
    
    On Error GoTo ErrHandler
    
    Set oOutware = New PlusLib2.COutWare
    oOutware.Connection = g_adoCon
    
    Set rs = oOutware.GetOrderSubTotal(grdOut.TextMatrix(grdOut.Row, 1))

    Set oOutware = Nothing
    
    With grdOutColor
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        Do Until rs.EOF
            .AddItem Format(rs!OrderSeq, "00") & vbTab & rs!Color & vbTab & rs!DesignNO & vbTab & rs!ColorQty & vbTab & "0" & vbTab & rs!OutQty
            
            nTotal = nTotal + CLng(rs!OutQty)
            rs.MoveNext
        Loop
        rs.Close
        
        .Redraw = flexRDDirect
    End With
   
    grdOutColorSum.TextArray(2) = nTotal
    
    Call CalcColorQty
    Call ChangeScroll(3)
    
    Set rs = Nothing
    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oOutware = Nothing
    Call ErrorBox(Err.Number, "frmOutwareWork.FillGridOutColor", Err.Description)
End Sub


Private Sub CalcColorQty()
    Dim i%, j%
    Dim nQtySum&, nQtyTotalSum&
    
    nQtySum = 0
    nQtyTotalSum = CLng(grdOutColorSum.TextArray(2))
      
    With grdOutColor
        For i = 1 To grdOut.Rows - 1
            For j = 1 To grdOutColor.Rows - 1
                If grdOut.TextMatrix(i, 2) = .TextMatrix(j, 0) Then
                    .TextMatrix(j, 4) = CInt(.TextMatrix(j, 4)) + CInt(grdOut.TextMatrix(i, 6))
                    .TextMatrix(j, 5) = CLng(.TextMatrix(j, 5)) + CInt(grdOut.TextMatrix(i, 6))
                    nQtySum = nQtySum + CInt(grdOut.TextMatrix(i, 6))
                    nQtyTotalSum = nQtyTotalSum + CInt(grdOut.TextMatrix(i, 6))
                    Exit For
                End If
            Next j
        Next i
    End With

    With grdOutColorSum
        .TextArray(1) = nQtySum
        .TextArray(2) = nQtyTotalSum
    End With
End Sub

Private Sub ChangeScroll(Index As Integer)
    Select Case Index
        Case 0
            With grdOrder
                .ColWidth(3) = LIMIT_WIDTH1 - IIf(.Rows > LIMIT_ROW1, 240, 0)
            End With
        Case 1
            With grdColor
                .ColWidth(1) = LIMIT_WIDTH2 - IIf(.Rows > LIMIT_ROW2, 240, 0)
            End With
        Case 2
            With grdOut
                .ColWidth(8) = LIMIT_WIDTH3 - IIf(.Rows > LIMIT_ROW3, 240, 0)
            End With
        Case 3
            With grdOutColor
                .ColWidth(1) = LIMIT_WIDTH4 - IIf(.Rows > LIMIT_ROW4, 240, 0)
            End With
    End Select
End Sub

Private Sub ClearData()
    Dim i%
    
    Call ClearText(txtName)
    
    For i = 5 To 7
        txtName(i) = 0
    Next i
    
End Sub

Private Sub ShowData()
    With grdOrder
        If optOrder(0) Then
            txtName(0) = .TextMatrix(.Row, 2)
        Else
            txtName(0) = .TextMatrix(.Row, 1)
        End If
        txtName(1) = .TextMatrix(.Row, 3)
        txtName(2) = .TextMatrix(.Row, 6)
        txtName(3) = .TextMatrix(.Row, 7)
        txtName(4) = .TextMatrix(.Row, 9)
        txtName(5) = Format(.TextMatrix(.Row, 8), "#,##0") & IIf(.TextMatrix(.Row, 10) = "0", " Y", " M")
        txtName(6) = Format(.TextMatrix(.Row, 5), "#,##0")
        txtName(7) = Format(.TextMatrix(.Row, 4), "#,##0")
        txtName(8) = Format(.TextMatrix(.Row, 11), "#,##0")
        txtName(9) = Format(.TextMatrix(.Row, 12), "#,##0")
    End With
    
    Call FillGridColor
End Sub

Private Function CheckRcv() As Boolean
    
    If comOut.InBufferCount > 0 Then
        CheckRcv = True
    Else
        CheckRcv = False
    End If
End Function

Private Function RcvFrame() As Integer
    Dim cnt As Long, nDatalen%
    Dim RcvHead As String
    Dim ch As String * 1, DataCnt%

    ' STX
    cnt = 0
    Do While cnt < 300000
        If CheckRcv Then
            ch = comOut.Input
            If ch = STX Then Exit Do
        Else
            cnt = cnt + 1
        End If
    Loop
    If cnt >= 300000 Then
        RcvFrame = -1
        Exit Function
    End If
    
    ' CMD(2) + LEN(4)
    cnt = 0
    DataCnt = 0
    Do While cnt < 300000
        If CheckRcv Then
            DataCnt = DataCnt + 1
            RcvHead = RcvHead & comOut.Input
            If DataCnt = 6 Then Exit Do
        Else
            cnt = cnt + 1
        End If
    Loop
    If cnt >= 300000 Then
        RcvFrame = -1
        Exit Function
    End If
    
    m_sRcvBuf = RcvHead
    nDatalen = val(Right(RcvHead, 4))
    
    If nDatalen <> 0 Then
        ' Data
        cnt = 0
        DataCnt = 0
        Do While cnt < 300000
            If CheckRcv Then
                DataCnt = DataCnt + 1
                m_sRcvBuf = m_sRcvBuf & comOut.Input
                If DataCnt = nDatalen Then Exit Do
            Else
                cnt = cnt + 1
            End If
        Loop
        If cnt >= 300000 Then
            RcvFrame = -1
            Exit Function
        End If
    End If
    
    ' ETX
    cnt = 0
    Do While cnt < 300000
        If CheckRcv Then
            ch = comOut.Input
            Exit Do
        Else
            cnt = cnt + 1
        End If
    Loop
    If cnt >= 300000 Or ch <> ETX Then
        comOut.Output = NAK
        RcvFrame = -1
        Exit Function
    Else
        comOut.Output = ACK
        RcvFrame = 1
    End If
    
End Function

Private Sub Sleep(val As Integer)
    Dim i%, j%
    Dim fValue As Single
    
    For i = 0 To val
        For j = 0 To 10000
            fValue = i / 8 * j
        Next j
    Next i
End Sub

Private Sub InitComm()
    Call EndComm
    With comOut
        .CommPort = cboCom.ListIndex + 1
        .Settings = "19200,n,8,1"
        .RTSEnable = True
        .RThreshold = 1
        .InputLen = 1
        .PortOpen = True
    End With
End Sub

Private Sub EndComm()
        If comOut.PortOpen Then comOut.PortOpen = False
End Sub

