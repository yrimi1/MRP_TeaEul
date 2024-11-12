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
      Caption         =   " [  계획현황  ]"
      Height          =   4410
      Left            =   30
      TabIndex        =   4
      Top             =   45
      Width           =   15165
      Begin VB.CommandButton cmdCheck 
         Caption         =   "전체 선택"
         Height          =   315
         Index           =   0
         Left            =   1305
         TabIndex        =   28
         Top             =   4050
         Width           =   1140
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "선택 해제"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   27
         Top             =   4050
         Width           =   1140
      End
      Begin VB.ComboBox cboProcessID 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1485
         Style           =   2  '드롭다운 목록
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
            Name            =   "굴림"
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
         Caption         =   "공정"
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
         Caption         =   "계획일자"
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
         Caption         =   "가져오기"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame fraKey 
      Caption         =   " [  계획작성  ]"
      Height          =   1605
      Left            =   30
      TabIndex        =   1
      Top             =   4560
      Width           =   15165
      Begin VB.TextBox txtOrderID 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
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
         Alignment       =   1  '오른쪽 맞춤
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
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   5025
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   570
         Width           =   1335
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "저장(&S)"
         Height          =   795
         Index           =   3
         Left            =   11130
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   10
         ToolTipText     =   "자료 저장"
         Top             =   720
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   795
         Index           =   0
         Left            =   12720
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   9
         ToolTipText     =   "자료 추가"
         Top             =   720
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   795
         Index           =   2
         Left            =   14310
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   8
         ToolTipText     =   "자료 삭제"
         Top             =   720
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   795
         Index           =   1
         Left            =   13515
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   7
         ToolTipText     =   "자료 수정"
         Top             =   720
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "취소(&C)"
         Height          =   795
         Index           =   4
         Left            =   11925
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   6
         ToolTipText     =   "자료 취소"
         Top             =   720
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.ComboBox cboEmerClss 
         Height          =   300
         ItemData        =   "frmPlanCPB.frx":000C
         Left            =   3675
         List            =   "frmPlanCPB.frx":000E
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   570
         Width           =   1305
      End
      Begin VB.ComboBox cboPlanClss 
         Height          =   300
         ItemData        =   "frmPlanCPB.frx":0010
         Left            =   2325
         List            =   "frmPlanCPB.frx":0012
         Style           =   2  '드롭다운 목록
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
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "긴급구분"
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
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "계획구분"
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
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "작성자"
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
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "내용"
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
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "색상명"
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
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "수량"
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
            Caption         =   "관리 번호"
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
         Caption         =   "관리 번호"
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
      Caption         =   "      닫기(&X)"
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
      Begin VB.CheckBox chkExpand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "공정확장"
         Height          =   500
         Left            =   0
         Style           =   1  '그래픽
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
Private m_nSelected As Integer ' 수주 선택갯수
'Private m_bSkipEvent As Boolean
Private m_bLoading As Boolean
Private m_iFlag    As Integer   ' 현재 상태 (추가/수정/삭제/검색)
Private m_ProcessID As String
'---------------------------------------------------------------

Private Sub cboProcessID_Click()
    
''    cboProcessID.AddItem "C염색"  '4000   칼라무  0
''    cboProcessID.AddItem "염색"   '4300   칼라유  1
    
    Call InitGrid
    Call ClearData
    txtOrderID.Text = ""
    If cboProcessID.ListIndex = 0 Then    'C염색
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
        
'    If NewValue Then '[1] 조회모드 = True
'        grdPlanData.Editable = flexEDNone
'    Else '[2] 편집모드 = False
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
        MsgBox ("금일 이전의 데이터는 등록, 수정, 삭제가 불가능 합니다.")
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
                MsgBox "수주번호를 먼저 입력 하십시오.", vbInformation
                Call SetCancel
            End If
            
            Select Case m_ProcessID
                Case "4000"  'c염색 -> 칼라 있음.
                    With grdOrder
                        If .Rows > .FixedRows Then
                            If .IsSubtotal(grdOrder.Row) = True Then
                                MsgBox "하위 내용을 선택하십시오", vbInformation
                                
                                'cancel과 같은 처리
                                Call SetCancel
                            Else
                                txtColorName.Text = grdOrder.TextMatrix(grdOrder.Row, 3)
                                txtColorName.Tag = GetOrderSeq(txtOrderID, txtColorName)
                            End If
                         End If
                    End With
                Case "4300"  '염색  ->
                    txtColorName.Text = ""
                    txtColorName.Tag = ""
            End Select
        
        '-------------------------------------------------------------------------------------'
        Case ID_UPDATE
            '레코드가 없을 경우
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
            nNoPlanQty = rs!ColorQtyYDS * (1 + rs!ChunkRate / 100) - (rs!InstQty - rs!배색Qty + rs!배색TQty)    '미계획량
            nProceTotalQty = rs!전처리Qty + rs!효소호발Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty + rs!수세Qty + _
                            rs!SKQty + rs!셋팅Qty + rs!PeachQty + rs!C염색Qty + rs!염색Qty + rs!P염색Qty + rs!R수세Qty + _
                            rs!건조Qty + rs!가공Qty + rs!검사Qty + rs!PauseQty
                            
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
                rs!InstQty - rs!배색Qty + rs!배색TQty & vbTab & rs!InstQty - rs!배색Qty & vbTab & rs!배색TQty & vbTab & nProceTotalQty & vbTab & _
                rs!전처리Qty + rs!효소호발Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty & vbTab & _
                rs!수세Qty & vbTab & rs!SKQty & vbTab & _
                rs!셋팅Qty & vbTab & rs!PeachQty & vbTab & rs!C염색Qty & vbTab & _
                rs!염색Qty + rs!P염색Qty + rs!R수세Qty & vbTab & rs!건조Qty & vbTab & rs!가공Qty & vbTab & _
                rs!검사Qty & vbTab & rs!PauseQty & vbTab & rs!PassQty & vbTab & rs!DefectQty & vbTab & rs!OutQty & vbTab & CheckNull(rs!PatternID) & vbTab & rs!WorkWidth
        
'            .TextMatrix(nTop, 7) = CLng(.TextMatrix(nTop, 7)) + rs!StuffInQty
            .TextMatrix(nTop, 8) = CLng(.TextMatrix(nTop, 8)) + nNoPlanQty
            .TextMatrix(nTop, 9) = CLng(.TextMatrix(nTop, 9)) + rs!InstQty - rs!배색Qty + rs!배색TQty
            .TextMatrix(nTop, 10) = CLng(.TextMatrix(nTop, 10)) + rs!InstQty - rs!배색Qty
            .TextMatrix(nTop, 11) = CLng(.TextMatrix(nTop, 11)) + rs!배색TQty
            .TextMatrix(nTop, 12) = CLng(.TextMatrix(nTop, 12)) + nProceTotalQty
            .TextMatrix(nTop, 13) = CLng(.TextMatrix(nTop, 13)) + rs!전처리Qty + rs!효소호발Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty
            .TextMatrix(nTop, 14) = CLng(.TextMatrix(nTop, 14)) + rs!수세Qty
            .TextMatrix(nTop, 15) = CLng(.TextMatrix(nTop, 15)) + rs!SKQty
            .TextMatrix(nTop, 16) = CLng(.TextMatrix(nTop, 16)) + rs!셋팅Qty
            .TextMatrix(nTop, 17) = CLng(.TextMatrix(nTop, 17)) + rs!PeachQty
            .TextMatrix(nTop, 18) = CLng(.TextMatrix(nTop, 18)) + rs!C염색Qty
            .TextMatrix(nTop, 19) = CLng(.TextMatrix(nTop, 19)) + rs!염색Qty + rs!P염색Qty + rs!R수세Qty
            .TextMatrix(nTop, 20) = CLng(.TextMatrix(nTop, 20)) + rs!건조Qty
            .TextMatrix(nTop, 21) = CLng(.TextMatrix(nTop, 21)) + rs!가공Qty
            .TextMatrix(nTop, 22) = CLng(.TextMatrix(nTop, 22)) + rs!검사Qty
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
        .AddItem "보통"
        .AddItem "긴급"
        .ListIndex = 0
    End With
    
    With cboPlanClss
        .Clear
        .AddItem "지급"
        .AddItem "수정"
        .ListIndex = 0
    End With

'    m_bLoading = False
    
    Call SetDtpDate(2, dtpPlanDate, dtpPlanDate)
    
    '-- 필수입력 항목에 icon 설정 하기

    pnlName(5).Picture = LoadResPicture("BASIC", vbResIcon)
    cmdLeft.Picture = LoadResPicture("LEFT", vbResIcon)
    cmdRight.Picture = LoadResPicture("RIGHT", vbResIcon)
    
    '공정선택 콤보박스에 C염색(C.P.B염색), 염색(Rapid염색)으로 설정하는 프로시저 호출
    'Call SetProcessID(cboProcessID, "'4000', '4300'")
    
    cboProcessID.AddItem "C염색"  '4000
    cboProcessID.AddItem "염색"   '4300
    cboProcessID.ListIndex = 0
    
'    Call FillGrdPlanCPB
    
    Call NonEditMode(True)
    
    txtOrderID.SetFocus
    
End Sub
Private Sub LoadPlanCPB(ByVal pProcID As Integer)
    cboProcessID.ListIndex = pProcID  '0: 4000(c염색), 1:Rapid 염색으로 설정
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
        
        .sPlanDate = MakeDate(DF_SHORT, dtpPlanDate)          '[2] 계획일자
        .sProcessID = IIf(cboProcessID.ListIndex = 0, "4000", "4300")             '[3] 공정코드
        .sOrderID = Trim(txtOrderID.Text)                     '[4] 관리번호
        .sPlanClss = Trim(cboPlanClss.Text)                   '[5] 지급, 수정
        .sEmerClss = Trim(cboEmerClss.Text)                   '[6] 긴급, 보통
        .sPersonID = g_sUserName                              '[7] 작성자 코드
        .sRemark = IIf(Len(Trim(txtRemark.Text)) = 0, " ", Trim(txtRemark.Text))                     '[8] 계획내역
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
            nNoPlanQty = rs!ColorQtyYDS * (1 + rs!ChunkRate / 100) - (rs!InstQty - rs!배색Qty + rs!배색TQty)    '미계획량
            
            nProceTotalQty = rs!전처리Qty + rs!호발Qty + rs!정련Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty + rs!셋팅Qty + rs!폭줄Qty + _
                            rs!PeachQty + rs!C염색Qty + rs!염색Qty + rs!P염색Qty + rs!R수세Qty + rs!건조Qty + rs!가공Qty + _
                            rs!검사Qty + rs!PauseQty
                            
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
                rs!InstQty - rs!배색Qty + rs!배색TQty & vbTab & rs!InstQty - rs!배색Qty & vbTab & rs!배색TQty & vbTab & nProceTotalQty & vbTab & _
                rs!전처리Qty + rs!호발Qty + rs!정련Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty & vbTab & _
                rs!셋팅Qty + rs!폭줄Qty & vbTab & rs!PeachQty & vbTab & rs!C염색Qty & vbTab & _
                rs!염색Qty + rs!P염색Qty + rs!R수세Qty & vbTab & rs!건조Qty & vbTab & rs!가공Qty & vbTab & _
                rs!검사Qty & vbTab & rs!PauseQty & vbTab & rs!PassQty & vbTab & rs!DefectQty & vbTab & rs!OutQty & vbTab & CheckNull(rs!PatternID)
        
'            .TextMatrix(nTop, 7) = CLng(.TextMatrix(nTop, 7)) + rs!StuffInQty
            .TextMatrix(nTop, 8) = CLng(.TextMatrix(nTop, 8)) + nNoPlanQty
            .TextMatrix(nTop, 9) = CLng(.TextMatrix(nTop, 9)) + rs!InstQty - rs!배색Qty + rs!배색TQty
            .TextMatrix(nTop, 10) = CLng(.TextMatrix(nTop, 10)) + rs!InstQty - rs!배색Qty
            .TextMatrix(nTop, 11) = CLng(.TextMatrix(nTop, 11)) + rs!배색TQty
            .TextMatrix(nTop, 12) = CLng(.TextMatrix(nTop, 12)) + nProceTotalQty
            .TextMatrix(nTop, 13) = CLng(.TextMatrix(nTop, 13)) + rs!전처리Qty + rs!호발Qty + rs!정련Qty + rs!정련FQty + rs!정련SQty + rs!감량SQty
            .TextMatrix(nTop, 14) = CLng(.TextMatrix(nTop, 14)) + rs!셋팅Qty + rs!폭줄Qty
            .TextMatrix(nTop, 15) = CLng(.TextMatrix(nTop, 15)) + rs!PeachQty
            .TextMatrix(nTop, 16) = CLng(.TextMatrix(nTop, 16)) + rs!C염색Qty
            .TextMatrix(nTop, 17) = CLng(.TextMatrix(nTop, 17)) + rs!염색Qty + rs!P염색Qty + rs!R수세Qty
            .TextMatrix(nTop, 18) = CLng(.TextMatrix(nTop, 18)) + rs!건조Qty
            .TextMatrix(nTop, 19) = CLng(.TextMatrix(nTop, 19)) + rs!가공Qty
            .TextMatrix(nTop, 20) = CLng(.TextMatrix(nTop, 20)) + rs!검사Qty
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

    '수주현황및 공정현황
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
        .TextArray(1) = "거래처":       .ColWidth(1) = 1750:            .ColAlignment(1) = flexAlignLeftCenter
        .TextArray(2) = "품명":         .ColWidth(2) = 1550:             .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "관리번호" & vbCrLf & "색  상  명":     .ColWidth(3) = 1350:            .ColAlignment(3) = flexAlignLeftCenter
        .TextArray(4) = "Order No." & vbCrLf & "색  상  명":   .ColWidth(4) = 0:               .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "축율":         .ColWidth(5) = 800:            .ColAlignment(5) = flexAlignCenterCenter
        .TextArray(6) = "수주량":       .ColWidth(6) = 1000:            .ColAlignment(6) = flexAlignRightCenter
        .TextArray(7) = "입고량":       .ColWidth(7) = 900:             .ColAlignment(7) = flexAlignRightCenter
        .TextArray(8) = "미계획량":     .ColWidth(8) = 900:             .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "계획량":       .ColWidth(9) = 900:             .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "배색":        .ColWidth(10) = 900:            .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "배색":        .ColWidth(11) = 900:            .ColAlignment(11) = flexAlignRightCenter
        .TextArray(12) = "공정량":      .ColWidth(12) = 900:            .ColAlignment(12) = flexAlignRightCenter
        .TextArray(13) = "정련":        .ColWidth(13) = 900:            .ColAlignment(13) = flexAlignRightCenter
        .TextArray(14) = "수세":        .ColWidth(14) = 900:            .ColAlignment(14) = flexAlignRightCenter
        .TextArray(15) = "S/K":         .ColWidth(15) = 900:            .ColAlignment(15) = flexAlignRightCenter
        .TextArray(16) = "SETT":        .ColWidth(16) = 900:            .ColAlignment(16) = flexAlignRightCenter
        .TextArray(17) = "PEACH":       .ColWidth(17) = 900:            .ColAlignment(17) = flexAlignRightCenter
        .TextArray(18) = "CPB":         .ColWidth(18) = 900:            .ColAlignment(18) = flexAlignRightCenter
        .TextArray(19) = "염색":        .ColWidth(19) = 900:            .ColAlignment(19) = flexAlignRightCenter
        .TextArray(20) = "DRY":         .ColWidth(20) = 900:            .ColAlignment(20) = flexAlignRightCenter
        .TextArray(21) = "가공":        .ColWidth(21) = 900:            .ColAlignment(21) = flexAlignRightCenter
        .TextArray(22) = "검사":        .ColWidth(22) = 900:            .ColAlignment(22) = flexAlignRightCenter
        .TextArray(23) = "보류":        .ColWidth(23) = 900:            .ColAlignment(23) = flexAlignRightCenter
        .TextArray(24) = "검사":        .ColWidth(24) = 900:            .ColAlignment(24) = flexAlignRightCenter
        .TextArray(25) = "검사":        .ColWidth(25) = 900:            .ColAlignment(25) = flexAlignRightCenter
        .TextArray(26) = "출고량":      .ColWidth(26) = 1000:           .ColAlignment(26) = flexAlignRightCenter
        .TextArray(27) = "공정패턴코드":        .ColWidth(27) = 0
        .TextArray(28) = "가공폭":      .ColWidth(28) = 0
        
        .TextArray(.Cols + 0) = " "
        .TextArray(.Cols + 1) = "거래처"
        .TextArray(.Cols + 2) = "품명"
        .TextArray(.Cols + 3) = "관리번호" & vbCrLf & "색  상  명"
        .TextArray(.Cols + 4) = "Order No." & vbCrLf & "색  상  명"
        .TextArray(.Cols + 5) = "축율"
        .TextArray(.Cols + 6) = "수주량"
        .TextArray(.Cols + 7) = "입고량"
        .TextArray(.Cols + 8) = "미계획량"
        .TextArray(.Cols + 9) = "계획량"
        .TextArray(.Cols + 10) = "대기량"
        .TextArray(.Cols + 11) = "배색량"
        .TextArray(.Cols + 12) = "공정량"
        .TextArray(.Cols + 13) = "정련"
        .TextArray(.Cols + 14) = "수세"
        .TextArray(.Cols + 15) = "S/K"
        .TextArray(.Cols + 16) = "SETT"
        .TextArray(.Cols + 17) = "PEACH"
        .TextArray(.Cols + 18) = "CPB"
        .TextArray(.Cols + 19) = "염색"
        .TextArray(.Cols + 20) = "DRY"
        .TextArray(.Cols + 21) = "가공"
        .TextArray(.Cols + 22) = "검사"
        .TextArray(.Cols + 23) = "보류"
        .TextArray(.Cols + 24) = "합격"
        .TextArray(.Cols + 25) = "불합격"
        .TextArray(.Cols + 26) = "출고량"
        .TextArray(.Cols + 27) = "공정패턴코드"
        .TextArray(.Cols + 28) = "가공폭"

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

    '긴급구분, 계획구분, 거래처, 품명, 오더번호, 관리번호, 공정명, 내역
    Call SetVSFlexGrid(grdPlanData)
    With grdPlanData
        .Redraw = flexRDNone

        .Row = 0
        .Cols = 11

        .TextArray(1) = "선택":       .ColWidth(1) = 500:              .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "계획":       .ColWidth(2) = 500:              .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "긴급":       .ColWidth(3) = 500:              .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "품명":       .ColWidth(4) = 2600:             .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "관리번호":   .ColWidth(5) = 1300:             .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "OrderID":    .ColWidth(6) = 1300:             .ColAlignment(6) = flexAlignLeftCenter
        .TextArray(7) = "색상명":     .ColWidth(7) = 2400:             .ColAlignment(7) = flexAlignLeftCenter
        .TextArray(8) = "수량":       .ColWidth(8) = 1000:             .ColAlignment(8) = flexAlignRightCenter
        .TextArray(9) = "내역":       .ColWidth(9) = 2000:             .ColAlignment(9) = flexAlignLeftCenter
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
        .TextArray(0) = "합계":           .ColWidth(0) = 2000:   .ColAlignment(0) = flexAlignCenterCenter
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
            Checked = IIf(.Cell(flexcpChecked, .Row, .Col) = flexChecked, False, True)  '체크되면 true, 체크해제는 false
            .Cell(flexcpChecked, .Row, .Col) = Checked
        End If
        optOrder(1).Value = True
        Call ShowData(MakeOrderID(.TextMatrix(.Row, 5), OM_REDUCE), .TextMatrix(.Row, 6), .TextMatrix(.Row, 10))
    End With
    
End Sub

'/********************************************************
' * Description : CPB / Rapid 생지 투입계획
' * 기       능 : pl_mast 계획일자 조건으로  select
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    날 짜        작성자    버전                   변경사항
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     최현숙    1.0     작성
' ********************************************************/
' 긴급구분, 계획구분, 거래처 , 품명, 오더번호, 관리번호, 공정명, 내역
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
' * Description : CPB / Rapid 생지 투입계획
' * 기       능 : pl_mast 계획일자 조건으로  select
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    날 짜        작성자    버전                   변경사항
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     최현숙    1.0     작성
' ********************************************************/
' 긴급구분, 계획구분, 거래처 , 품명, 오더번호, 관리번호, 공정명, 내역
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
            '-- 진행과정 Progress Bar표시
            
            '-- 데이터 grid에 display
            .AddItem CStr(i) & vbTab & vbTab & rs!PlanClss & vbTab & rs!EmerClss & vbTab & _
                rs!ArticleName & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                Trim(rs!OrderNo) & vbTab & Trim(rs!Color) & vbTab & _
                SetCurrency(rs!Qty, 0) & vbTab & rs!Remark & vbTab & rs!OrderSeq
                
                nTotQty = nTotQty + CheckNum(rs!Qty)
                
            '-- 2줄마다 칼라 넣기
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
'            Checked = IIf(.Cell(flexcpChecked, .Row, .Col) = flexChecked, False, True)  '체크되면 true, 체크해제는 false
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

