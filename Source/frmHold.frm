VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHold 
   Caption         =   "보류"
   ClientHeight    =   9255
   ClientLeft      =   180
   ClientTop       =   735
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15180
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   706
      TabCaption(0)   =   "보류중"
      TabPicture(0)   =   "frmHold.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSPanel4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pnlProc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grdHold"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "대기중"
      TabPicture(1)   =   "frmHold.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grdWait"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "pnlHold"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "SSPanel2(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboProcID(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   -74910
         TabIndex        =   27
         Top             =   480
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   1455
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton cmdSearch 
            Caption         =   "검색(&F)"
            Height          =   690
            Left            =   13335
            MousePointer    =   99  '사용자 정의
            Style           =   1  '그래픽
            TabIndex        =   43
            ToolTipText     =   "자료 검색"
            Top             =   60
            Width           =   1605
         End
         Begin VB.ComboBox cboProcID 
            Height          =   300
            Index           =   0
            Left            =   7590
            Style           =   2  '드롭다운 목록
            TabIndex        =   30
            Top             =   450
            Width           =   1695
         End
         Begin VB.TextBox txtOrderID 
            Height          =   300
            Left            =   4545
            TabIndex        =   29
            Top             =   450
            Width           =   1500
         End
         Begin VB.TextBox txtCardID 
            Height          =   300
            Left            =   7575
            TabIndex        =   28
            Top             =   90
            Width           =   1500
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   6
            Left            =   3120
            TabIndex        =   31
            Top             =   450
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   529
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkOrderID 
               Caption         =   "관리번호"
               Height          =   240
               Left            =   60
               TabIndex        =   32
               Top             =   45
               Width           =   1200
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   0
            Left            =   6360
            TabIndex        =   33
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkCardID 
               Caption         =   "카드번호"
               Height          =   240
               Left            =   60
               TabIndex        =   34
               Top             =   45
               Width           =   1080
            End
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   1
            Left            =   1305
            TabIndex        =   35
            Top             =   90
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
            Left            =   1305
            TabIndex        =   36
            Top             =   435
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   36871
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   37
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkDate 
               Caption         =   "보류일자"
               Height          =   240
               Left            =   60
               TabIndex        =   38
               Top             =   45
               Value           =   1  '확인
               Width           =   1080
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   6360
            TabIndex        =   39
            Top             =   450
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkProcID 
               Caption         =   "공정"
               Height          =   240
               Left            =   60
               TabIndex        =   40
               Top             =   30
               Width           =   1080
            End
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   345
            Left            =   3120
            TabIndex        =   44
            Top             =   60
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   609
            _Version        =   196609
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.OptionButton optOrder 
               Caption         =   "관리 번호"
               Height          =   180
               Index           =   1
               Left            =   1380
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   90
               Value           =   -1  'True
               Width           =   1110
            End
            Begin VB.OptionButton optOrder 
               Caption         =   "Order No."
               Height          =   180
               Index           =   0
               Left            =   60
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   90
               Width           =   1170
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   3
            Left            =   10020
            TabIndex        =   50
            Top             =   90
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   529
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkProc 
               Caption         =   "처리포함"
               Height          =   210
               Left            =   60
               TabIndex        =   51
               Top             =   60
               Width           =   1080
            End
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            Caption         =   "부터"
            Height          =   180
            Index           =   3
            Left            =   2580
            TabIndex        =   42
            Top             =   150
            Width           =   360
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            Caption         =   "까지"
            Height          =   180
            Index           =   2
            Left            =   2580
            TabIndex        =   41
            Top             =   495
            Width           =   360
         End
      End
      Begin VB.ComboBox cboProcID 
         Height          =   300
         Index           =   1
         Left            =   1410
         Style           =   2  '드롭다운 목록
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin Threed.SSPanel pnlProc 
         Height          =   795
         Left            =   -74910
         TabIndex        =   1
         Top             =   7650
         Width           =   14970
         _ExtentX        =   26405
         _ExtentY        =   1402
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton cmdProc 
            Caption         =   "처리방안 작성"
            Enabled         =   0   'False
            Height          =   690
            Left            =   13335
            MousePointer    =   99  '사용자 정의
            Style           =   1  '그래픽
            TabIndex        =   47
            ToolTipText     =   "자료 검색"
            Top             =   60
            Width           =   1605
         End
         Begin VB.TextBox txtProcOpinion 
            Height          =   645
            Left            =   5010
            TabIndex        =   7
            Top             =   60
            Width           =   4005
         End
         Begin VB.TextBox txtProcPerson 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            TabIndex        =   6
            Top             =   405
            Width           =   1785
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Index           =   0
            Left            =   3690
            TabIndex        =   3
            Top             =   60
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "처리방안"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   300
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   405
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "처리자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   300
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   60
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "처리일시"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpProcDate 
            Height          =   300
            Left            =   1470
            TabIndex        =   8
            Top             =   60
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   36871
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   690
            Left            =   11640
            TabIndex        =   49
            Top             =   60
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   1217
            _Version        =   196609
            Caption         =   "      인쇄(&P)"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "대기공정"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlHold 
         Height          =   1575
         Left            =   90
         TabIndex        =   11
         Top             =   6870
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   2778
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CommandButton cmdHold 
            Caption         =   "보류 처리"
            Height          =   675
            Left            =   13335
            MousePointer    =   99  '사용자 정의
            Style           =   1  '그래픽
            TabIndex        =   48
            ToolTipText     =   "자료 검색"
            Top             =   870
            Width           =   1590
         End
         Begin VB.TextBox txtMainHold 
            Height          =   315
            Left            =   10980
            TabIndex        =   23
            Top             =   450
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.ComboBox cboProcID 
            Height          =   300
            Index           =   2
            Left            =   1470
            Style           =   2  '드롭다운 목록
            TabIndex        =   18
            Top             =   810
            Width           =   1695
         End
         Begin VB.TextBox txtPersonName 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1470
            TabIndex        =   13
            Top             =   450
            Width           =   1785
         End
         Begin VB.TextBox txtHoldReason 
            Height          =   1335
            Left            =   5160
            TabIndex        =   12
            Top             =   90
            Width           =   4785
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Index           =   4
            Left            =   3840
            TabIndex        =   14
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "보류원인"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Index           =   5
            Left            =   150
            TabIndex        =   15
            Top             =   450
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "작성자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Index           =   6
            Left            =   150
            TabIndex        =   16
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "보류일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   3
            Left            =   1470
            TabIndex        =   17
            Top             =   90
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   36871
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Index           =   7
            Left            =   150
            TabIndex        =   19
            Top             =   810
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "발생공정"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Index           =   8
            Left            =   150
            TabIndex        =   20
            Top             =   1140
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "발생일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   4
            Left            =   1470
            TabIndex        =   21
            Top             =   1140
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   36871
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Index           =   9
            Left            =   10980
            TabIndex        =   22
            Top             =   90
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "주보류명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   300
            Left            =   12390
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   90
            Visible         =   0   'False
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
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
      End
      Begin VSFlex7LCtl.VSFlexGrid grdHold 
         Height          =   6255
         Left            =   -74910
         TabIndex        =   25
         Top             =   1380
         Width           =   14985
         _cx             =   26432
         _cy             =   11033
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
      Begin VSFlex7LCtl.VSFlexGrid grdWait 
         Height          =   6015
         Left            =   90
         TabIndex        =   26
         Top             =   810
         Width           =   14985
         _cx             =   26432
         _cy             =   10610
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
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   13530
      TabIndex        =   2
      Top             =   8550
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmHold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TParaHold
    nCheckDate      As Integer
    sDate           As String
    eDate           As String
    nCheckOrderID   As Integer
    OrderID         As String
    nCheckOrderNo   As Integer
    OrderNo         As String
    nCheckCardID    As Integer
    CardID          As String
    SplitID         As String
    nCheckProcID    As Integer
    WriteProcID     As String
    nCheckProcClss  As Integer
End Type

'-------------------------------------------------
Private Type TParaHoldRec
    nAffected    As Integer
    sJobFlag     As String
    WriteDate    As String
    WriteProcID  As String
    WriteSeq     As Integer
    CardID       As String
    SplitID      As String
    WorkSeq      As Integer
    OrderID      As String
    OrderSeq     As Integer
    PersonID     As String
    OccuProcID   As String
    OccuDate     As String
    MainHold     As String
    HoldReason   As String
End Type
'------------------------------------------------
Private Type TOpinion
    nAffected    As Integer
    WriteDate    As String
    WriteProcID  As String
    WriteSeq     As Integer
    ProcPerson   As String
    ProcDate     As String
    ProcOpinion  As String
    CardID       As String
    SplitID      As String
End Type

Private Sub InitGrid()
    Dim II%
    
    Call SetVSFlexGrid(grdHold)
    With grdHold
        .Cols = 17
        
        .Redraw = flexRDNone

        .Rows = 4
        .FixedRows = 4
        .FixedCols = 1
    '    .FrozenCols = 3
        .RowHeight(0) = 400

        .TextMatrix(3, 0) = "처리":                        .ColWidth(0) = 400
        .TextMatrix(3, 1) = "보류일자":                    .ColWidth(1) = 1000:            .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(3, 2) = "카드번호":                    .ColWidth(2) = 1400:            .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(3, 3) = "관리번호":                    .ColWidth(3) = 1300:            .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(3, 4) = "OrderNo":                     .ColWidth(4) = 1300:            .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(3, 5) = "거래처":                      .ColWidth(5) = 1200:            .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(3, 6) = "품명":                        .ColWidth(6) = 1800:            .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(3, 7) = "색상명":                      .ColWidth(7) = 1800:            .ColAlignment(7) = flexAlignLeftCenter
        .TextMatrix(3, 8) = "수량":                        .ColWidth(8) = 800:             .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(3, 9) = "절수":                        .ColWidth(9) = 600:             .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(3, 10) = "보류공정":                   .ColWidth(10) = 1100:           .ColAlignment(10) = flexAlignCenterCenter
        .TextMatrix(3, 11) = "보류원인":                   .ColWidth(11) = 3000:           .ColAlignment(11) = flexAlignLeftCenter
        .TextMatrix(3, 12) = "보류" & vbCrLf & "작성자":   .ColWidth(12) = 800:            .ColAlignment(12) = flexAlignCenterCenter
        .TextMatrix(3, 13) = "처리방안":                   .ColWidth(13) = 3000:           .ColAlignment(13) = flexAlignLeftCenter
        .TextMatrix(3, 14) = "방안" & vbCrLf & "작성자":   .ColWidth(14) = 800:            .ColAlignment(14) = flexAlignCenterCenter
        .TextMatrix(3, 15) = "방안" & vbCrLf & "작성일":   .ColWidth(15) = 800:           .ColAlignment(15) = flexAlignCenterCenter
        .TextMatrix(3, 16) = "WORKsEQ":        .ColWidth(16) = 0: .ColHidden(16) = True
        
        For II = 0 To .Cols - 1
            .FixedAlignment(II) = flexAlignCenterCenter
        Next II
        
        .ColHidden(9) = True
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        
        .RowHeight(0) = 400
        .RowHeight(1) = 400
        .RowHeight(2) = 400
        .RowHeight(3) = 400

        .MergeCells = flexMergeFree
        
        .AllowUserResizing = flexResizeBoth
        .Redraw = flexRDDirect
    End With

    Call SetVSFlexGrid(grdWait)
    With grdWait
        .Cols = 14

        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1
        .FixedCols = 1
        .FrozenCols = 3
        .RowHeight(0) = 400

        .TextArray(0) = " ":               .ColWidth(0) = 20
        .TextArray(1) = "대기공정":        .ColWidth(1) = 1100:            .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "카드번호":        .ColWidth(2) = 1400:            .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "관리번호":        .ColWidth(3) = 1300:            .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "OrderNo":         .ColWidth(4) = 1300:            .ColAlignment(4) = flexAlignCenterCenter
        .TextArray(5) = "거래처":          .ColWidth(5) = 1200:            .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "품명":            .ColWidth(6) = 1800:            .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "색상명":          .ColWidth(7) = 1800:             .ColAlignment(7) = flexAlignLeftCenter
        .TextArray(8) = "절수":            .ColWidth(8) = 600:             .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "수량":            .ColWidth(9) = 1000:             .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "후공정":         .ColWidth(10) = 10000:           .ColAlignment(10) = flexAlignLeftCenter
        
        .TextArray(11) = "CardID"
        .TextArray(12) = "WorkSeq"
        .TextArray(13) = "OrderSeq"
        
        .ColHidden(11) = True
        .ColHidden(12) = True
        .ColHidden(13) = True
    
        .AllowUserResizing = flexResizeBoth
        .Redraw = flexRDDirect
    End With
    
End Sub



Private Sub cboProcID_Click(Index As Integer)
    Dim dProcID As String
    If Trim(cboProcID(Index)) <> AllStr Then
        dProcID = GetProcessID(cboProcID(Index))
        cboProcID(Index).Tag = dProcID
    Else
        cboProcID(Index).Tag = ""
    End If
    
    Select Case Index
        Case 1
            Call FillgrdWait
    End Select
End Sub

Private Sub chkCardID_Click()
    If chkDate.Value = vbChecked Then
        txtCardID.Enabled = True
        txtCardID.SetFocus
    Else
        txtCardID.Enabled = False
    End If
    
End Sub

Private Sub chkDate_Click()
    If chkDate.Value = vbChecked Then
        dtpDate(1).Enabled = True
        dtpDate(2).Enabled = True
    Else
        dtpDate(1).Enabled = False
        dtpDate(2).Enabled = False
    End If
End Sub

Private Sub chkOrderID_Click()
    If chkOrderID.Value = vbChecked Then
        txtOrderID.Enabled = True
        txtOrderID.SetFocus
    Else
        txtOrderID.Enabled = False
    End If
End Sub

Private Sub chkProcID_Click()
    If chkDate.Value = vbChecked Then
        cboProcID(0).Enabled = True
        cboProcID(0).ListIndex = 0
    Else
        cboProcID(0).Enabled = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub FillgrdHold()
    Dim adoCmd As ADODB.Command
    Dim dRS As ADODB.Recordset
    Dim TParaHold As TParaHold
    
    cmdProc.Enabled = False
    '------ Parameter 넘겨줄 값 Move
    With TParaHold
        If chkOrderID.Value = vbChecked Then
            If optOrder(0).Value = True Then  'Order NO
                .nCheckOrderID = 0
                
                .OrderID = ""
                
                .nCheckOrderNo = 1
                .OrderNo = txtOrderID.Text
            Else
                .nCheckOrderID = 1
                .OrderID = txtOrderID.Text
                
                .nCheckOrderNo = 0
                .OrderNo = ""
            End If
        Else
            .nCheckOrderID = 0
            .OrderID = ""
            .nCheckOrderNo = 0
            .OrderNo = ""
        End If
        
        If chkDate.Value = vbChecked Then
            .nCheckDate = 1
            .sDate = MakeDate(DF_SHORT, dtpDate(1))
            .eDate = MakeDate(DF_SHORT, dtpDate(2))
        Else
            .nCheckDate = 0
            .sDate = ""
            .eDate = ""
        End If
        
        If chkCardID.Value = vbChecked Then
            .nCheckCardID = 1
            .CardID = Left(txtCardID.Text, 8)
            .SplitID = Mid(txtCardID.Text, 9)
        Else
            .nCheckCardID = 0
            .CardID = ""
            .SplitID = ""
        End If
        
        If chkProcID.Value = vbChecked Then
            .nCheckProcID = 1
            .WriteProcID = GetProcessID(Trim(cboProcID(0).Text))
        Else
            .nCheckProcID = 0
            .WriteProcID = ""
        End If
        
        If chkProc.Value = vbChecked Then
            .nCheckProcClss = 0     '처리포함 ( ALL )
        Else
            .nCheckProcClss = 1     '미처리만
        End If
    End With
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Hold_sHoldDraft"
        
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, TParaHold.nCheckDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TParaHold.sDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TParaHold.eDate)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, TParaHold.nCheckOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, TParaHold.OrderID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, TParaHold.nCheckOrderNo)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 24, TParaHold.OrderNo)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, TParaHold.nCheckCardID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TParaHold.CardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, TParaHold.SplitID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, TParaHold.nCheckProcID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, TParaHold.WriteProcID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, TParaHold.nCheckProcClss)
        
        Set dRS = .Execute
    End With
    
    Set adoCmd = Nothing
    
  '  Call SetVSFlexGrid(grdHold)
    With grdHold
        .Rows = .FixedRows
        .Redraw = flexRDNone
        .ExplorerBar = flexExNone
        .ScrollBars = flexScrollBarBoth
        
        If .Rows > .FixedRows Then
            .Row = 1
        End If

        Do Until dRS.EOF
            
            .AddItem IIf(Trim(dRS!ProcPersonName) = "", "", "■") & vbTab & MakeDate(DF_LONG, dRS!WriteDate) & vbTab & _
                        MakeCardID(dRS!CardID, OM_EXPAND, dRS!SplitID) & vbTab & _
                        MakeOrderID(dRS!OrderID, OM_EXPAND) & vbTab & _
                        Trim(dRS!OrderNo) & vbTab & Trim(dRS!kCustom) & vbTab & Trim(dRS!Article) & vbTab & _
                        Trim(dRS!ColorName) & vbTab & dRS!Qty & vbTab & dRS!Roll & vbTab & Trim(dRS!WriteProcID) & vbTab & Trim(dRS!HoldReason) & vbTab & _
                        Trim(dRS!PersonName) & vbTab & Trim(dRS!ProcOpinion) & vbTab & Trim(dRS!ProcPersonName) & vbTab & _
                        Format(dRS!ProcDate, "MM/DD HH:NN") & vbTab & dRS!WriteSeq
            
            
            .RowHeight(.Rows - 1) = 350
            dRS.MoveNext
        Loop
        dRS.Close
        Set dRS = Nothing
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        Else
            MsgBox LoadResString(203), vbInformation
        End If
    End With
End Sub




Private Sub cmdHold_Click()
    Dim TRec As TParaHoldRec
    Dim adoCmd As ADODB.Command
    Dim dRS As ADODB.Recordset
    Dim TParaHold As TParaHold
    Dim irow As Integer, nSql As Integer
    
    If grdWait.Rows = grdWait.FixedRows Then
        Exit Sub
    Else
        irow = grdWait.Row
    End If
    
    If Trim(txtHoldReason) = "" Then
        MsgBox "보류 원인을 입력하지 않았습니다", vbInformation, "보류 처리"
        Exit Sub
    End If
        
    With TRec
        .nAffected = 0
        .sJobFlag = "I"
        .WriteDate = Format$(Now, "yyyymmdd")
        .WriteProcID = GetProcessID(grdWait.TextMatrix(irow, 1))
        .WorkSeq = 0
        .CardID = Left(grdWait.TextMatrix(irow, 11), 8)
        .SplitID = Mid(grdWait.TextMatrix(irow, 11), 9)
        .WorkSeq = grdWait.TextMatrix(irow, 12)
        .OrderID = MakeOrderID(grdWait.TextMatrix(irow, 3), OM_REDUCE)
        .OrderSeq = grdWait.TextMatrix(irow, 13)
        .PersonID = txtPersonName.Tag
        .OccuProcID = GetProcessID(Trim(cboProcID(2).Text))
        .OccuDate = MakeDate(DF_SHORT, dtpDate(4))
        .MainHold = ""
        .HoldReason = Trim(txtHoldReason.Text)
    End With
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Hold_iuHold"
        
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamOutput, 1, TRec.nAffected)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, TRec.sJobFlag)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TRec.WriteDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, TRec.WriteProcID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInputOutput, 1, TRec.WriteSeq)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TRec.CardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, TRec.SplitID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, TRec.WorkSeq)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, TRec.OrderID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, TRec.OrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TRec.PersonID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, TRec.OccuProcID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TRec.OccuDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 3, TRec.MainHold)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 100, TRec.HoldReason)
        
        .Execute
    End With
    Set adoCmd = Nothing
    Call ClearPnlHold
    Call FillgrdHold
    Call FillgrdWait

End Sub

Private Sub cmdPrint_Click()

    If MsgBox("인쇄 하시겠습니까?", vbYesNo) = vbYes Then
        Call ColResize("-")
        Call FillGrdPrint
        Call ColResize("+")
    End If
    
End Sub
Sub ColResize(ByVal pType As String)
    Dim II%
    
    If pType = "-" Then
        With grdHold
            For II = 0 To .Cols - 1
            .ColWidth(II) = Int(.ColWidth(II) * 0.7)
           Next II
            .Redraw = flexRDDirect
        End With
    Else
        With grdHold
            For II = 0 To .Cols - 1
            .ColWidth(II) = Int(.ColWidth(II) / 0.7)
           Next II
            .Redraw = flexRDDirect
        End With
    End If
    
    

End Sub

Private Sub cmdProc_Click()
    Dim TRec As TParaHoldRec
    Dim adoCmd As ADODB.Command
    Dim dRS As ADODB.Recordset
    Dim TOpinion As TOpinion
    Dim irow As Integer, nSql As Integer
    
    If grdHold.Rows = grdHold.FixedRows Then
        Exit Sub
    Else
        irow = grdHold.Row
    End If
    
    If Trim(txtProcOpinion) = "" Then
        MsgBox "처리 방안을 입력하지 않았습니다", vbInformation, "처리방안 작성"
        Exit Sub
    End If
    
    With TOpinion
        .nAffected = 0
        .WriteDate = MakeDate(DF_SHORT, grdHold.TextMatrix(irow, 1))
        .WriteProcID = GetProcessID(grdHold.TextMatrix(irow, 10))
        .WriteSeq = grdHold.TextMatrix(irow, 16)
        .ProcOpinion = Trim(txtProcOpinion.Text)
        .ProcDate = MakeDate(DF_SHORT, dtpProcDate)
        .ProcPerson = g_sUserName
        .CardID = Left(MakeCardID(grdHold.TextMatrix(irow, 2), OM_REDUCE), 8)
        .SplitID = Mid(MakeCardID(grdHold.TextMatrix(irow, 2), OM_REDUCE), 9)
    End With
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Hold_uProcOpinion"
        
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamOutput, 1, TOpinion.nAffected)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TOpinion.WriteDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, TOpinion.WriteProcID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, TOpinion.WriteSeq)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TOpinion.ProcPerson)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TOpinion.ProcDate)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 100, TOpinion.ProcOpinion)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TOpinion.CardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, TOpinion.SplitID)
        
        .Execute
        
         nSql = .Parameters(0).Value
    End With
    Set adoCmd = Nothing
    
    Call FillgrdHold
    Call FillgrdWait
    Call ClearPnlProc

End Sub

Private Sub cmdSearch_Click()
    Call FillgrdHold
End Sub
Sub FillgrdWait()
    Dim adoCmd As ADODB.Command
    Dim dRS As ADODB.Recordset
    Dim TParaHold As TParaHold
    Dim nCheckProc As Integer
    
    
    nCheckProc = 0
    If Trim(cboProcID(1).Tag) <> "" Then
        nCheckProc = 1
    End If
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Hold_sWaitDraft"
        
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 1, nCheckProc)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, Trim(cboProcID(1).Tag))
        
        Set dRS = .Execute
    End With
    Set adoCmd = Nothing
    
    
    Call SetVSFlexGrid(grdWait)
    
    With grdWait
        .Rows = .FixedRows
        .Redraw = flexRDNone
        .ExplorerBar = flexExNone
        .ScrollBars = flexScrollBarBoth
        
        If .Rows > .FixedRows Then
            .Rows = 1
        End If

        Do Until dRS.EOF
            
            .AddItem .Rows & vbTab & Trim(dRS!WaitProcName) & vbTab & MakeCardID(dRS!CardID, OM_EXPAND, dRS!SplitID) & vbTab & _
                         MakeOrderID(dRS!OrderID, OM_EXPAND) & vbTab & _
                        Trim(dRS!OrderNo) & vbTab & Trim(dRS!kCustom) & vbTab & Trim(dRS!Article) & vbTab & _
                        Trim(dRS!ColorName) & vbTab & dRS!Roll & vbTab & dRS!Qty & vbTab & Trim(dRS!AfterProc) & vbTab & _
                        Trim(dRS!CardID) & Trim(dRS!SplitID) & vbTab & dRS!WorkSeq & vbTab & dRS!OrderSeq
            
            .RowHeight(.Rows - 1) = 350
            dRS.MoveNext
        Loop
        dRS.Close
        Set dRS = Nothing
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If
    End With


End Sub



Private Sub ClearPnlProc()
    dtpProcDate = Now
    txtProcOpinion = ""
    txtProcPerson.Text = g_sPersonName
    txtProcPerson.Tag = g_sUserName
    cmdProc.Enabled = False
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    Dim i%
    
    PlusMDI.pnlMenu.Visible = False
    
    Me.Move 0, 0, 15300, 9660
    
    Call InitGrid
    Call SetOperate(Me)
    
    
    dtpProcDate = Now
    dtpDate(1) = Now
    dtpDate(2) = Now
    dtpDate(3) = Now
    dtpDate(4) = Now
    
    Call SetComboProcss(cboProcID(0), AllStr)
    Call SetComboProcss(cboProcID(1), AllStr)
    Call SetComboProcss(cboProcID(2))
    
'    dtpDate(1).Enabled = False
'    dtpDate(2).Enabled = False
    
    txtCardID.Enabled = False
    txtOrderID.Enabled = False
    cboProcID(0).Enabled = False
    
    Call FillgrdWait
    Call ClearPnlProc
    Call ClearPnlHold
    
    
'    txtPersonName.Text = g_sUserName
'    txtPersonName.Tag = g_sPersonName
End Sub
Sub FillGrdPrint()
    Dim i%
    Dim sDate As String, eDate As String, sOrderID As String
    
    '---- 보류일자
    If chkDate.Value Then
        sDate = Format(dtpDate(1), "YYYY/MM/DD")
        eDate = Format(dtpDate(2), "YYYY/MM/DD")
    Else
        sDate = ""
        eDate = ""
    End If
    
    If chkOrderID.Value Then
        If optOrder(0).Value = True Then
            sOrderID = "Order NO: " & Trim(txtOrderID)
        Else
            sOrderID = "관리번호: " & Trim(txtOrderID)
        End If
    Else
        sOrderID = "OrderNO: (전체) "
    End If
    
    With grdHold
        .Redraw = flexRDBuffered
        .FrozenCols = 0

        .GridLinesFixed = flexGridFlat
        .GridColorFixed = vbWhite
        .GridLines = flexGridFlat
        .RowHidden(0) = False
        .RowHidden(1) = False
        .RowHidden(2) = False
        .RowHeight(0) = 500
        .RowHeight(1) = 500
        .RowHeight(2) = 500
        
        .FontSize = 7
        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "보류현황"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        
        .Cell(flexcpText, 1, 1, 1, 4) = "▶ 보류일자 : " & sDate & " ~ " & eDate
        .Cell(flexcpText, 1, 5, 1, 6) = "▶ " & Trim(sOrderID)
        .Cell(flexcpText, 1, .Cols - 5, 1, .Cols - 4) = "▶ 발행일 : " & Format(Now, "YYYY/MM/DD hh:mm")
        
        .Cell(flexcpText, 1, 7, 1, 8) = "▶ 카드번호 : " & IIf(chkCardID.Value, Trim(txtCardID.Text), "(전체)")
        .Cell(flexcpText, 2, 1, 2, 2) = "▶ 공    정 : " & IIf(chkProcID.Value, Trim(cboProcID(1).Text), "(전체)")
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite

        For i = 0 To 3
           .MergeRow(i) = True
        Next i
        
        .PrintGrid "태을염직", True, 2, 0, 500
        .GridLinesFixed = flexGridInset
        
        .Redraw = flexRDDirect
        
        .FontSize = 9

        For i = 0 To 2
            .RowHidden(i) = True
        Next i
        
        .Redraw = flexRDDirect
    End With
    
End Sub


Private Sub ClearPnlHold()
    dtpDate(3) = Now
    txtPersonName.Text = g_sPersonName
    txtPersonName.Tag = g_sUserName
    cboProcID(2).ListIndex = 0
    dtpDate(4) = Now
    dtpProcDate = Now
    txtHoldReason.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub




Private Sub grdHold_RowColChange()
    With grdHold
        Call ClearPnlProc
        If .TextMatrix(.Row, 0) = "" Then
            cmdProc.Enabled = True
        Else
            cmdProc.Enabled = False
            dtpProcDate = CDate(.TextMatrix(.Row, 15))
            txtProcOpinion = .TextMatrix(.Row, 13)
            txtProcPerson.Text = .TextMatrix(.Row, 14)
            txtProcPerson.Tag = .TextMatrix(.Row, 14)
        End If
    End With

End Sub

Private Sub optOrder_Click(Index As Integer)
    chkOrderID.Caption = optOrder(Index).Caption
    
End Sub


''Private Sub SSCommand2_Click()
''    Select Case Index
''    Case 3
''        Call ReturnCode(LG_ORDER, 0, True, txtOrder)
''    Case 4
''        Call ReturnCode(LG_CUSTOM, 0, True, txtCustomID)
''    Case 5
''        Call ReturnCode(LG_ARTICLE, 0, True, txtArticleID)
''    End Select
''
''End Sub
Private Sub SSPanel4_Click()

End Sub

