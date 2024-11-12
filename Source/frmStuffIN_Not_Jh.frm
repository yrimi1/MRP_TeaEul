VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffIN_Not_JH 
   Caption         =   "생지 입고 관리"
   ClientHeight    =   9390
   ClientLeft      =   2745
   ClientTop       =   1890
   ClientWidth     =   11910
   Icon            =   "frmStuffIN_Not_Jh.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   11910
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10080
      TabIndex        =   14
      Top             =   8520
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin TabDlg.SSTab tabForm 
      Height          =   8835
      Left            =   45
      TabIndex        =   13
      Top             =   30
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   15584
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   741
      TabCaption(0)   =   "    조 회   "
      TabPicture(0)   =   "frmStuffIN_Not_Jh.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pnlCaption(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dtpDate(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dtpDate(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdFind(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pnlCaption(11)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "pnlCaption(9)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdFind(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "grdGroup"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtArticle"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCustom(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdOperate(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdOperate(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdOperate(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optGroup(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "optGroup(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdTerm(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdTerm(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "SSPanel1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "CboStuffClss2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdSearch"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "SSPanel2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "   입력화면   "
      TabPicture(1)   =   "frmStuffIN_Not_Jh.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pnlData"
      Tab(1).ControlCount=   1
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Left            =   60
         TabIndex        =   54
         Top             =   780
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   609
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   3
            Left            =   1500
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   90
            Width           =   1140
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   90
            Value           =   -1  'True
            Width           =   1140
         End
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검색(&F)"
         Height          =   780
         Left            =   8100
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   52
         ToolTipText     =   "자료 저장"
         Top             =   90
         Width           =   780
      End
      Begin VB.ComboBox CboStuffClss2 
         Height          =   300
         Left            =   5730
         Style           =   2  '드롭다운 목록
         TabIndex        =   47
         Top             =   780
         Width           =   1965
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   4410
         TabIndex        =   46
         Top             =   780
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "입고구분"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   48
            Top             =   60
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "전일"
         Height          =   315
         Index           =   0
         Left            =   1110
         MousePointer    =   99  '사용자 정의
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   450
         Width           =   615
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1110
         MousePointer    =   99  '사용자 정의
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   90
         Width           =   615
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "거래처별"
         Height          =   300
         Index           =   1
         Left            =   60
         Style           =   1  '그래픽
         TabIndex        =   36
         Top             =   465
         Width           =   990
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "오더별"
         Height          =   300
         Index           =   0
         Left            =   60
         Style           =   1  '그래픽
         TabIndex        =   35
         Top             =   105
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "수정(&U)"
         Height          =   780
         Index           =   1
         Left            =   10080
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   34
         ToolTipText     =   "자료 수정"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "삭제(&D)"
         Height          =   780
         Index           =   2
         Left            =   10875
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   33
         ToolTipText     =   "자료 삭제"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "추가(&A)"
         Height          =   780
         Index           =   0
         Left            =   9285
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   32
         ToolTipText     =   "자료 추가"
         Top             =   60
         Width           =   780
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         Index           =   1
         Left            =   5730
         TabIndex        =   25
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtArticle 
         Height          =   300
         Left            =   5730
         TabIndex        =   24
         Top             =   450
         Width           =   1935
      End
      Begin VSFlex7LCtl.VSFlexGrid grdGroup 
         Height          =   7170
         Left            =   60
         TabIndex        =   23
         Top             =   1140
         Width           =   11700
         _cx             =   20637
         _cy             =   12647
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
      Begin Threed.SSPanel pnlData 
         Height          =   8280
         Left            =   -74940
         TabIndex        =   15
         Top             =   60
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   14605
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtOrderNO 
            Height          =   315
            Left            =   5475
            TabIndex        =   60
            Top             =   720
            Width           =   2340
         End
         Begin VB.TextBox txtSearch 
            Height          =   315
            Left            =   1350
            TabIndex        =   58
            Top             =   720
            Width           =   2340
         End
         Begin VB.CommandButton Command1 
            Caption         =   "다른색상입력"
            Height          =   345
            Left            =   75
            TabIndex        =   53
            Top             =   2745
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.ComboBox cboName 
            Height          =   300
            Index           =   14
            Left            =   9780
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   1440
            Width           =   1365
         End
         Begin VB.TextBox TxtArticleID2 
            Height          =   300
            IMEMode         =   10  '한글 
            Left            =   5475
            TabIndex        =   6
            Top             =   1440
            Width           =   2340
         End
         Begin VB.TextBox txtThreadName 
            Height          =   300
            IMEMode         =   10  '한글 
            Left            =   5475
            TabIndex        =   5
            Top             =   1095
            Width           =   2340
         End
         Begin VB.TextBox txtCustomID 
            Height          =   300
            IMEMode         =   10  '한글 
            Left            =   1350
            TabIndex        =   2
            Top             =   1095
            Width           =   2355
         End
         Begin VB.TextBox txtStuffSeq 
            BackColor       =   &H00E0E0E0&
            Height          =   345
            Left            =   8115
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   135
            Width           =   1005
         End
         Begin VB.ComboBox CboStuffClss 
            Height          =   300
            Left            =   4695
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   165
            Width           =   2175
         End
         Begin VB.TextBox txtNum 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   7725
            TabIndex        =   19
            Top             =   2775
            Width           =   1170
         End
         Begin VB.TextBox txtNum 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   10455
            TabIndex        =   21
            Top             =   2760
            Width           =   1170
         End
         Begin VB.TextBox txtRemark 
            Height          =   630
            Left            =   1290
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   11
            Top             =   7560
            Width           =   8535
         End
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   9780
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   1095
            Width           =   1350
         End
         Begin VB.TextBox txtCustom 
            Height          =   300
            IMEMode         =   10  '한글 
            Index           =   0
            Left            =   1350
            TabIndex        =   4
            Top             =   1440
            Width           =   2340
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   315
            Index           =   2
            Left            =   1275
            TabIndex        =   0
            Top             =   165
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyy-MM-dd (ddd)"
            Format          =   23789571
            CurrentDate     =   37068
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   10
            Left            =   3555
            TabIndex        =   16
            Top             =   165
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "입고 구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   3
            Left            =   60
            TabIndex        =   17
            Top             =   7560
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "비고 사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VSFlex7LCtl.VSFlexGrid grdStuffINSub 
            Height          =   4395
            Left            =   60
            TabIndex        =   10
            Top             =   3120
            Width           =   11565
            _cx             =   20399
            _cy             =   7752
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
            Height          =   300
            Index           =   5
            Left            =   8370
            TabIndex        =   22
            Top             =   1095
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "입고 단위"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdSave 
            Height          =   690
            Left            =   9990
            TabIndex        =   12
            Top             =   7545
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   1217
            _Version        =   196609
            Caption         =   "      저장(&S)"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   1
            Left            =   3720
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1095
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   6315
            TabIndex        =   18
            Top             =   2775
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "전체입고절수"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   9030
            TabIndex        =   20
            Top             =   2775
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "전체입고수량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   7
            Left            =   6975
            TabIndex        =   43
            Top             =   165
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "입고 순번"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   8
            Left            =   135
            TabIndex        =   44
            Top             =   165
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "입고 일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   13
            Left            =   4245
            TabIndex        =   49
            Top             =   1095
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "사종"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   4
            Left            =   7845
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1440
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   14
            Left            =   4245
            TabIndex        =   50
            Top             =   1440
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "품명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   12
            Left            =   8370
            TabIndex        =   51
            Top             =   1440
            Width           =   1365
            _ExtentX        =   2408
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
            Caption         =   "가공 구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   510
            Left            =   9315
            TabIndex        =   57
            Top             =   60
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   900
            _Version        =   196609
            BackColor       =   65535
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin VSFlex7LCtl.VSFlexGrid grdData 
            Height          =   780
            Left            =   30
            TabIndex        =   59
            Top             =   1830
            Width           =   11655
            _cx             =   20558
            _cy             =   1376
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
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   16777215
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
            Height          =   315
            Index           =   15
            Left            =   4245
            TabIndex        =   61
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "OrderNO"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   1440
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "입고처 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   12
            Left            =   120
            TabIndex        =   63
            Top             =   1095
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "거래처"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "관리번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   315
            Index           =   3
            Left            =   7830
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   720
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   2
            Index           =   1
            X1              =   0
            X2              =   11715
            Y1              =   2700
            Y2              =   2700
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   2
            Index           =   0
            X1              =   -30
            X2              =   11685
            Y1              =   615
            Y2              =   615
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   315
         Index           =   0
         Left            =   7710
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   9
         Left            =   4410
         TabIndex        =   27
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거 래 처"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   11
         Left            =   4410
         TabIndex        =   29
         Top             =   450
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "품     명"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   1065
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   315
         Index           =   2
         Left            =   7710
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   3120
         TabIndex        =   39
         Top             =   105
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23789569
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   3120
         TabIndex        =   40
         Top             =   450
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23789569
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   6
         Left            =   1800
         TabIndex        =   41
         Top             =   90
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "입고 일자"
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   42
            Top             =   60
            Value           =   1  '확인
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmStuffIN_Not_JH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_iFlag As String * 1
Dim m_bGroupClss As Boolean     '거래처별, 오더별 Grid 구분

Private Const LIMIT_WIDTH1 = 1640
Private Const LIMIT_WIDTH2 = 2100
Private Const LIMIT_WIDTH3 = 560
Private Const LIMIT_WIDTH4 = 2000
Private Const LIMIT_ROW1 = 11
Private Const LIMIT_ROW2 = 28
Private Const LIMIT_ROW3 = 9
Private m_bSortForward As Boolean
Private m_StuffDate As String, m_StuffClss As String, m_StuffSeq As Integer


Private Sub cboName_KeyPress(Index As Integer, KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub


Private Sub CboStuffClss_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub Command1_Click()
    With grdStuffINSub
        .Rows = .Rows + 1
        .Select .Rows - 1, 1
    End With
    grdStuffINSub.SetFocus
End Sub

Private Sub dtpDate_KeyPress(Index As Integer, KeyAscii As Integer)
    Call MoveFocus(KeyAscii)
End Sub

Private Sub Form_Load()
    Dim i%
    Me.Move 0, 0, 11970, 9660

    Call InitGrid
    Call InitGroup
    
    Call SetOperate(Me)
    
    '----- 입고단위 설정
    With cboUnit
        .AddItem "YDS"
        .AddItem "MTS"
        .ListIndex = 0
    End With
    
    '----- 입고구분 설정
    With CboStuffClss
        .AddItem "1.생지"
        .AddItem "2.Shortage변상분"
        .AddItem "3.반품 생지"
        .AddItem "4.반품 완제품(가공불량)"
    End With
    
    '----- 검색용 입고구분 설정
    With CboStuffClss2
        .AddItem "1.생지"
        .AddItem "2.Shortage변상분"
        .AddItem "3.반품 생지"
        .AddItem "4.반품 완제품(가공불량)"
    End With
    
    Call MakeCodeCombo(cboName(14), CD_WORK)        ' 가공 구분
    
    '---- 날짜 설정
    For i = 0 To 2
        dtpDate(i) = Now
    Next i
    
    cmdSave.MousePointer = ssCustom
    cmdSave.Picture = LoadResPicture("SAVE", vbResIcon)
    cmdSave.MouseIcon = LoadResPicture("POINTER", vbResCursor)
    
    '--- find 컨트롤 icon설정
    For i = 0 To cmdFind.Count - 1
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        cmdFind(i).MouseIcon = LoadResPicture("POINTER", vbResCursor)
    Next i

    cmdFind(1).Visible = True
    cmdFind(3).Visible = True
    
    cmdFind(0).Enabled = False
    cmdFind(2).Enabled = False
    txtArticle.Enabled = False
    txtSearch.Enabled = True

    m_iFlag = ID_ADDNEW
    
    Call ClearText(txtNum, "0")
    
    m_bGroupClss = True
    Call FillGridGroup(m_bGroupClss)

'    pnlMsg.Caption = LoadResString(121)
'    optOrder(0).Value = True
    CboStuffClss.ListIndex = 0
    Call optOrder_Click(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Call SaveSetting(LoadResString(100), Me.Name, "Order", IIf(chkSearch(0) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "Custom", IIf(chkSearch(1) = vbChecked, "1", "0"))
    Call SaveSetting(LoadResString(100), Me.Name, "Article", IIf(chkSearch(2) = vbChecked, "1", "0"))
End Sub

'--- 입력 Text Clear
Private Sub ClearData()
    txtSearch.Text = ""
    CboStuffClss.ListIndex = 0
    txtCustomID.Text = ""
    txtCustomID.Tag = ""
    txtStuffSeq.Text = ""
    txtCustom(0).Text = ""
    txtThreadName.Text = ""
    TxtArticleID2.Text = ""
    TxtArticleID2.Tag = ""
    txtNum(0).Text = ""
    txtNum(1).Text = ""
    cboUnit.ListIndex = 0
    cboName(14).ListIndex = 0
End Sub
    
Private Sub ClearGridSub()
    Dim i%, j%
    With grdStuffINSub
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        .Rows = .FixedRows + 1
    End With
End Sub

Private Function CheckData() As Boolean
    CheckData = True
    
    If val(txtNum(0)) = 0 Or val(txtNum(1)) = 0 Then
        MsgBox "입고 절수(또는 수량)이 없습니다. (절수)수량을 입력하십시오.", vbInformation
        CheckData = False
        Exit Function
    End If
End Function

Private Sub SetNewData(SetStuffINData As PlusLib2.TStuffIN)
    Dim sJobFlag As String
    
    On Error GoTo ErrHandler
    
    Select Case m_iFlag
        Case ID_ADDNEW
            sJobFlag = "I"
        Case ID_UPDATE
            sJobFlag = "U"
    End Select
    
    With SetStuffINData
        .sJobFlag = sJobFlag
        .sStuffDate = MakeDate(DF_SHORT, dtpDate(2))     '[2] 입고 일자
        .sStuffClss = Left(CboStuffClss, 1)              '입고구분
        .nStuffSeq = val(txtStuffSeq)                    '순번
        .sCustomID = txtCustomID.Tag                     '발주처코드
        .sCustom = Trim(txtCustom(0))                    '원단입고처명
        .nTotRoll = val(txtNum(0))                       '원단절수
        .nTotQty = val(txtNum(1))                        '원단수량
        .sRemark = Trim(txtRemark.Text)                  '비고
        .sThreadName = Trim(txtThreadName.Text)          '사종
        .sUnitClss = cboUnit.ListIndex                   '입고단위
'        .UnitClss = "1"                                 '입고단위
        .sOrderID = Trim(txtSearch.Text)                 '관리번호
        .sWorkID = Format(cboName(14).ItemData(cboName(14).ListIndex), "0000")       ' 가공 구분
        .sArticleID = TxtArticleID2.Tag                  'item
        .sOrderNo = txtOrderNO                           'OrderNo
    End With

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Description, "frmStuffIN.SetNewData", Err.Description)

End Sub

''Private Function CheckSub() As Integer
''    Dim nSeq%
''    Dim nCount%, nPosition%
''    Dim i%, j%, k%
''
''    With grdStuffINSub
''        If .Rows = .FixedRows Then
''            CheckSub = -1
''            Exit Function
''        End If
''
''        For i = 1 To .Rows - 1
''            For j = 1 To .Cols - 1
''                If Len(.TextMatrix(i, j)) <> 0 Then
''                    nPosition = InStr(.TextMatrix(i, j), "*")
''                    If nPosition > 0 Then
''                        nCount = Mid(.TextMatrix(i, j), nPosition + 1)
''                        nSeq = nSeq + nCount
''                    Else
''                        nSeq = nSeq + 1
''                    End If
''                End If
''            Next j
''        Next i
''    End With
''    CheckSub = nSeq
''End Function

Private Sub SetNewDataSub(SetSubData() As PlusLib2.TStuffINSub, nSeq As Integer)
    Dim iLoop%, nCount%
    Dim i%, j%, k%
    Dim nRollQty%  '--- 1절당 수량
    Dim nQty%      '--- 전체수량 절수 * 수량
    
    Dim nStuffQty() As Integer      '수량을 넣는 배열
    Dim nStuColor() As String       'Color를 넣는 배열
    
    If nSeq >= 0 Then
        ReDim nStuffQty(nSeq)
        ReDim nStuColor(nSeq)
    End If
    
    
    With grdStuffINSub
        If .Rows = .FixedRows Then Exit Sub
        
        For i = 1 To .Rows - 1
            
            For j = 2 To .Cols - 1
                If Len(.TextMatrix(i, j)) <> 0 Then
                    
                    '--- 절수 및 수량 가져오기
                    Call GetRollQty(.TextMatrix(i, j), nCount, nRollQty, nQty)
                    
                    If nCount > 1 Then
                        For k = 0 To nCount - 1
                            nStuffQty(iLoop) = nRollQty
                            nStuColor(iLoop) = .TextMatrix(i, 1)
                            iLoop = iLoop + 1
                        Next k
                    Else
                        '--- 색상명 배열에 넣기
                        nStuColor(iLoop) = .TextMatrix(i, 1)
                        nStuffQty(iLoop) = nRollQty
                        iLoop = iLoop + 1
                    End If
                End If
            Next j
        Next i
    End With
    
    ' 입고세부 구조체 배열에 데이터 넣기
    For iLoop = 0 To nSeq
        With SetSubData(iLoop)
            .sColor = nStuColor(iLoop)         '칼라명
            .sColorID = IIf(Len(nStuColor(iLoop)) > 0, GetColorID(nStuColor(iLoop)), "")        '칼라ID
            .nQty = nStuffQty(iLoop)            '수량
        End With
    Next iLoop
    
    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffIN.SetNewDataSub", Err.Description)

End Sub
Function GetColorID(ByVal ColorName As String) As String
    Dim dSql_str As String
    Dim dRS As New ADODB.Recordset
    
    dSql_str = "SELECT ColorID FROM mt_color " & vbCr & _
               " WHERE Color = '" & ColorName & "' "
    dRS.Open dSql_str, g_adoCon, adOpenStatic, adLockReadOnly
    
    If dRS.RecordCount > 0 Then
        GetColorID = Trim$(dRS(0))
    End If
    dRS.Close
    Set dRS = Nothing
                       
End Function
Private Sub CalcQty()
    Dim i%, j%
    Dim nRoll%, nQty%, nRollQty%       '합계
    Dim nTotRoll%, nTotQty%
    
    ' Grid Text 값 계산
    With grdStuffINSub
        If .Rows = .FixedRows Then Exit Sub
        
        For i = 1 To .Rows - 1
            For j = 2 To .Cols - 1
                If Len(.TextMatrix(i, j)) <> 0 Then
                    Call GetRollQty(.TextMatrix(i, j), nRoll, nRollQty, nQty)
                    nTotRoll = nTotRoll + nRoll
                    nTotQty = nTotQty + nQty
                End If
            Next j
        Next i
    End With

    txtNum(0) = nTotRoll
    txtNum(1) = nTotQty
End Sub

Private Function DeleteData() As Boolean
    Dim oStuffIn As PlusLib2.CStuffIN

    On Error GoTo ErrHandler
    
    
'    Call FillGrid
   
    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    
    DeleteData = oStuffIn.DeleteStuffIN(m_StuffDate, m_StuffClss, m_StuffSeq)
    txtSearch.Tag = ""
    Set oStuffIn = Nothing
    Exit Function
    
ErrHandler:
    DeleteData = False
    Set oStuffIn = Nothing
    Call ErrorBox(Err.Number, "frmStuffIn.DeleteData", Err.Description)
End Function

Private Function SaveData() As Boolean
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim NewStuffIN   As PlusLib2.TStuffIN
    Dim StuffINSub() As PlusLib2.TStuffINSub
    Dim nSeq  As Integer
    Dim nStuffSeq As Integer
    
    On Error GoTo ErrHandler
    
    SaveData = False
    
    Dim sWorkID$
    
    '절수 개수 만큼 배열의 크기를 만든다.
    ' nSeq = CheckSub
    
    nSeq = val(txtNum(0))
    
    ' cStuffIn 클래스의 구조체에 값 대입
    Call SetNewData(NewStuffIN)
    
    '--- Stuffinsub 구조체에 값 넣기
    If nSeq > 0 Then
        ReDim StuffINSub(nSeq - 1)
        Call SetNewDataSub(StuffINSub, nSeq - 1)
    End If
    
    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    SaveData = oStuffIn.AddNewStuffIN(NewStuffIN, StuffINSub, nStuffSeq)
    txtStuffSeq = nStuffSeq
    
    Set oStuffIn = Nothing
    
    Exit Function

ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffIN.SaveData", Err.Description)
    Set oStuffIn = Nothing

End Function
Sub SetKeyEdit(ByVal dEdit As Boolean)
    dtpDate(2).Enabled = dEdit
    CboStuffClss.Enabled = dEdit
    txtStuffSeq.Enabled = dEdit
End Sub
'-----------------------------------------------------------------------
' StuffIN , StuffINSub Record 나타내기 및 StuffINSub Grid에 나타내기
' 수정 모드에서 사용함
'
'-------------- 입고데이터 나타내기  --- 수정모드
'-----------------------------------------------------------------------
Private Sub FillGridStuffSub(ByVal StuffDate As String, ByVal StuffClss As String, ByVal StuffSeq As Integer)
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim rs As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim iRow%, i%, j%, iCount%
    
    On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    
    ''''' StuffIN 1개의 record을 읽어온다.
    Set rs = oStuffIn.GetStuffINOne(StuffDate, StuffClss, StuffSeq)
    
    ''''' StuffINSub 의 데이터를 읽어 온다
    Set rsData = oStuffIn.GetStuffINSubONE(StuffDate, StuffClss, StuffSeq)
    
    Set oStuffIn = Nothing
    
    dtpDate(2) = MakeDate(DF_LONG, StuffDate)
    CboStuffClss.ListIndex = StuffClss - 1
    txtStuffSeq.Text = StuffSeq
    Call SetKeyEdit(False)
    
    With rsData
        txtCustomID.Tag = Trim$(rs!CustomID)
        txtCustomID.Text = Trim$(rs!kCustom)
        txtCustom(0).Text = Trim$(rs!Custom)
        txtThreadName.Text = Trim$(rs!ThreadName)
        TxtArticleID2.Text = Trim$(rs!Article)
        TxtArticleID2.Tag = Trim$(rs!ArticleID)
        txtNum(0).Text = Trim$(rs!TotRoll)
        txtNum(1).Text = Trim$(rs!TotQty)
        cboUnit = Trim$(rs!UnitName)
        cboName(14).ListIndex = FindItem(cboName(14), rs!WorkName)
        txtSearch.Text = Trim$(rs!OrderID)
    
    End With
    
    rs.Close
    Set rs = Nothing
    
    
    '''''' StuffSub 데이터 나타내기
    Dim nCols As Integer, nRows As Integer
    nCols = 1
    
    '--- grdStuffINSub grid 초기화
    grdStuffINSub.Rows = grdStuffINSub.FixedRows
    
    '---- 1번째 Color명과 같은지 확인 후 같은 라인(row)에 나타낸다.
    '---- 만약 다른경우 다음라인에 1번째 Col에 color를 나타낸후 표시 한다.
    Do Until rsData.EOF
    
        With grdStuffINSub
            .Redraw = flexRDNone
            
            '---맨처음 레코드 넣기
            If .Rows = .FixedRows Then
                .Rows = .Rows + 1
                nRows = .Rows - 1
                nCols = 1
                .TextMatrix(nRows, nCols) = CheckNull(rsData!Color)
                nCols = nCols + 1
            ElseIf Trim(.TextMatrix(nRows, 1)) <> CheckNull(rsData!Color) Or nCols >= 12 Then
                    .Rows = .Rows + 1
                    nRows = .Rows - 1
                    nCols = 1
                    .TextMatrix(nRows, nCols) = CheckNull(rsData!Color)
                    nCols = nCols + 1
            End If
            grdStuffINSub.TextMatrix(nRows, nCols) = rsData!Qty
            nCols = nCols + 1
            .Redraw = flexRDDirect
        End With
        rsData.MoveNext
    Loop
    
    
    
    Call ChangeScroll(2)
    rsData.Close
    Set rsData = Nothing
    
    '--- Order 에 대한 내용 grid에 나타내기
    Call FillStuffOrderData(txtSearch.Text)

    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffIN.FillGridStuffSub", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
End Sub


Private Sub chkSearch_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    '   Case 0
    '        If chkSearch(0) Then
    '            txtSearch.Enabled = True
    '            txtSearch.SetFocus
    '        Else
    '            txtSearch.Enabled = False
    '            cmdSearch.SetFocus
    '        End If
        Case 1    '거래처
            If chkSearch(1) = vbChecked Then
                txtCustom(1).Enabled = True
                txtCustom(1).SetFocus
                cmdFind(0).Enabled = True
            Else
                txtCustom(1).Enabled = False
                cmdFind(0).Enabled = False
                txtCustom(1).Tag = ""
            End If
        Case 2    '품명
            If chkSearch(2) = vbChecked Then
                txtArticle.Enabled = True
                txtArticle.SetFocus
                cmdFind(2).Enabled = True
            Else
                txtArticle.Enabled = False
                txtArticle.Tag = ""
                cmdSearch.SetFocus
                cmdFind(2).Enabled = False
            End If
        Case 3     '입고일자 Term
            If chkSearch(3) = vbChecked Then
                dtpDate(0).Enabled = True
                dtpDate(1).Enabled = True
            Else
                dtpDate(0).Enabled = False
                dtpDate(1).Enabled = False
            End If
        Case 4     '입고구분
            If chkSearch(Index) = vbChecked Then
                CboStuffClss2.Enabled = True
            Else
                CboStuffClss2.Enabled = False
            End If
            
            
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdData
        .Redraw = flexRDNone
        .Rows = 1
        .Cols = 14
        Call SetVSFlexGrid(grdData)
        
        .TextArray(1) = "관리번호":         .ColWidth(1) = 0:                   .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "Order NO":         .ColWidth(2) = 1300:                .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "접수일자":         .ColWidth(3) = 1050:                .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "거래처":           .ColWidth(4) = LIMIT_WIDTH1:        .ColAlignment(4) = flexAlignLeftCenter
        .TextArray(5) = "품명":             .ColWidth(5) = 1700:                .ColAlignment(5) = flexAlignLeftCenter
        .TextArray(6) = "구분":             .ColWidth(6) = 700:                 .ColAlignment(6) = flexAlignCenterCenter
        .TextArray(7) = "원단폭":           .ColWidth(7) = 660:                 .ColAlignment(7) = flexAlignCenterCenter
        .TextArray(8) = "축율" & vbCrLf & "LOSS":             .ColWidth(8) = 700:                 .ColAlignment(8) = flexAlignCenterCenter
        .TextArray(9) = "색상수":           .ColWidth(9) = 800:                 .ColAlignment(9) = flexAlignRightCenter
        .TextArray(10) = "주문량":          .ColWidth(10) = 1050:               .ColAlignment(10) = flexAlignRightCenter
        .TextArray(11) = "입고량":          .ColWidth(11) = 1050:               .ColAlignment(11) = flexAlignRightCenter
        .TextArray(12) = "주문" & vbCrLf & "단위": .ColWidth(12) = 600:         .ColAlignment(12) = flexAlignCenterCenter
        .TextArray(13) = "입고절수":        .ColWidth(13) = 0:                  .ColAlignment(13) = flexAlignCenterCenter
        
        .Redraw = flexRDDirect
    End With
    
    ' 절수, 수량 입력 Grid
    
    
    Dim sComBoList$
    
    
    '--- 사용가능한 ColorID, Color명 가져와서 Grid의 combobox의 item으로 설정
    Dim rs As ADODB.Recordset
    
    Set rs = GetColor
    sComBoList$ = " " & "|"
    Do Until rs.EOF
        sComBoList$ = sComBoList$ & rs(0) & "|"
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
  
    With grdStuffINSub
        .Redraw = flexRDNone
        .FixedRows = 1
        .FixedCols = 1
    Call SetVSFlexGrid(grdStuffINSub)
        .Cols = 12
        .Rows = .FixedRows + 1
        .TextArray(0) = ""
        .ColWidth(0) = LIMIT_WIDTH3
        
        '----------------------------------------------------------------------------------------------
        '------ 진호염직의 경우 Color관리를 하지 않기 때문에 2번째 Color Combo Col의 width를 0으로 했음.
        '------ 만약 다른업체에서 Color관리를 한다면 width = 1300으로 설정 함.
        '.ColComboList(1) = sComBoList
        '.ColWidth(1) = 1300
        '----------------------------------------------------------------------------------------------
        
        
        .ColWidth(1) = 0
       
        For i = 2 To .Cols - 1
            .TextArray(i) = i - 1
            .ColWidth(i) = Int((grdStuffINSub.Width - .ColWidth(0) - .ColWidth(1)) / 10)
        Next i
        
        .Editable = flexEDKbdMouse
        .FocusRect = flexFocusHeavy
        .Redraw = flexRDBuffered
        
        .ScrollBars = flexScrollBarBoth
        .Redraw = flexRDDirect

    End With
    

End Sub

Private Sub ChangeScroll(Index As Integer)
    Select Case Index
    Case 0
        With grdData
            If .Rows > LIMIT_ROW1 Then
                .ColWidth(4) = LIMIT_WIDTH1 - 240
            Else
                .ColWidth(4) = LIMIT_WIDTH1
            End If
        End With
    Case 1
        With grdGroup
            If m_bGroupClss Then
                If .Rows > LIMIT_ROW2 Then
                    .ColWidth(5) = LIMIT_WIDTH2 - 240
                Else
                    .ColWidth(5) = LIMIT_WIDTH2
                End If
            Else
                If .Rows > LIMIT_ROW2 Then
                    .ColWidth(7) = LIMIT_WIDTH4 - 240
                Else
                    .ColWidth(7) = LIMIT_WIDTH4
                End If
            
            End If
        End With
    Case 3
        With grdStuffINSub
            If .Rows > LIMIT_ROW3 Then
                .ColWidth(0) = LIMIT_WIDTH3 - 240
            Else
                .ColWidth(0) = LIMIT_WIDTH3
            End If
        End With
    End Select
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0                '[1] 거래처 코드
            Call ReturnCode(LG_CUSTOM, , False, txtCustom(1))
        Case 1                '[2] 거래처 코드
            Call ReturnCode(LG_CUSTOM, , False, txtCustomID)
        Case 2                '[3] 품명 코드
            Call ReturnCode(LG_ARTICLE, , False, txtArticle)
        Case 3                '[4] 오더 코드
            Call ReturnCode(LG_ORDER, , False, txtOrderNO)
            txtSearch.Text = txtOrderNO.Tag
            
        Case 4                '[4] 품명 코드
            Call ReturnCode(LG_ARTICLE, , False, TxtArticleID2)
    End Select
End Sub


Private Sub cmdOperate_Click(Index As Integer)
    Dim sStuffKey As String
    
    Select Case Index
        Case ID_ADDNEW
            m_iFlag = ID_ADDNEW
            tabForm.Tab = 1
            pnlMsg.Caption = "자료입력(추가)중..."
            Call ClearData
            Call SetKeyEdit(True)
            grdData.Rows = grdData.FixedRows
            grdStuffINSub.Rows = grdStuffINSub.FixedRows
            grdStuffINSub.Rows = grdStuffINSub.FixedRows + 1

        Case ID_UPDATE
            If grdGroup.Rows = grdGroup.FixedRows Then
                MsgBox LoadResString(203), vbInformation
                Exit Sub
            End If
            
            If grdGroup.IsSubtotal(grdGroup.Row) = True Then
                MsgBox "하위 내용을 선택하십시오", vbInformation
                Exit Sub
            End If
            
            pnlMsg.Caption = "자료입력(수정)중..."

            m_iFlag = ID_UPDATE
            
            '일자+구분+일련번호( StuffIN Key가져오기 )
            sStuffKey = grdGroup.TextMatrix(grdGroup.Row, grdGroup.Cols - 1)
        
            Call MakeStuffKey(sStuffKey, m_StuffDate, m_StuffClss, m_StuffSeq)
            
            Call FillGridStuffSub(m_StuffDate, m_StuffClss, m_StuffSeq)
            
            tabForm.Tab = 1

        Case ID_DELETE
            If grdGroup.Rows = grdGroup.FixedRows Then
                MsgBox LoadResString(203), vbInformation
                Exit Sub
            End If
            If grdGroup.IsSubtotal(grdGroup.Row) = True Then
                MsgBox "하위 내용을 선택하십시오", vbInformation
                Exit Sub
            End If
            
            '일자+구분+일련번호( StuffIN Key가져오기 )
            sStuffKey = grdGroup.TextMatrix(grdGroup.Row, grdGroup.Cols - 1)
        
            Call MakeStuffKey(sStuffKey, m_StuffDate, m_StuffClss, m_StuffSeq)
            
            tabForm.Tab = 1
            
            Call FillGridStuffSub(m_StuffDate, m_StuffClss, m_StuffSeq)
            
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "삭제확인") = vbYes Then
                m_iFlag = ID_DELETE
                If DeleteData Then
                    
                End If
            End If
            If optGroup(0) Then
                Call FillGridGroup
                Call optGroup_Click(0)
            Else
                Call FillGridGroup(False)
                Call optGroup_Click(1)
            End If
            
            m_iFlag = ""
            Call ClearData
            Call ClearGridSub
            grdData.Rows = grdData.FixedRows
            grdData.HighLight = flexHighlightNever
            tabForm.Tab = 0
    End Select

End Sub


Private Sub cmdSave_Click()
    If CheckData = False Then Exit Sub
    
    If SaveData Then
        MsgBox "입력한 내용이 저장 되었습니다.", vbInformation
        Call SetClearEdit
    End If
End Sub

Private Sub SetClearEdit()
    Call ClearData
    Call ClearGridSub
    grdData.Rows = grdData.FixedRows
    grdData.HighLight = flexHighlightNever
    txtSearch.Tag = ""
    txtSearch = ""
    m_iFlag = ID_ADDNEW
    Call SetKeyEdit(True)
    pnlMsg.Caption = "자료입력(추가)중..."
End Sub

Private Sub cmdSearch_Click()
    Dim Index As Integer
    
    Select Case tabForm.Tab
    Case 0
        If optGroup(0) Then
            Call FillGridGroup(True)
        Else
            Call InitGroup(False)
            Call FillGridGroup(False)
        End If
''    Case 1
''        Call FillGrid
    End Select
    Call optOrder_Click(Index)
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[3] 금일
        dtpDate(0) = Date - 1
        dtpDate(1) = Date - 1
    ElseIf Index = 1 Then '[2] 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    End If
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub SetGridGroup(NewFlex As VSFlexGrid)
    With NewFlex
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .GridLines = flexGridNone
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .BackColorBkg = vbWhite
        .SheetBorder = vbWhite
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarSimple
        .OutlineCol = 1
        .RowHeight(0) = 450
        .RowHeightMin = 275
    End With
End Sub

Private Sub InitGroup(Optional NewValue As Boolean = True)
    Dim i%
    Call SetGridGroup(grdGroup)
    If NewValue Then
        With grdGroup
            .Redraw = flexRDNone
            
            .FixedRows = 1
            .FixedCols = 0
            .Rows = 1
            .Cols = 21 ' 17
            
            .TextArray(0) = "":                                 .ColWidth(0) = 250
            .TextArray(1) = "관리번호":                         .ColAlignment(1) = flexAlignCenterCenter:
            .TextArray(2) = "Order NO":                         .ColAlignment(2) = flexAlignLeftCenter:
            .TextArray(3) = "접수일자":                         .ColAlignment(3) = flexAlignCenterCenter:
            .TextArray(4) = "거  래  처" & vbCrLf & "입고처":   .ColAlignment(4) = flexAlignLeftCenter:
            .TextArray(5) = "품명" & vbCrLf & "입고일자(사종)": .ColAlignment(5) = flexAlignLeftCenter:
            .TextArray(6) = "구분":                             .ColAlignment(6) = flexAlignCenterCenter:
            .TextArray(7) = "원단폭":                           .ColAlignment(7) = flexAlignRightCenter:
            .TextArray(8) = "축율" & vbCrLf & "LOSS":                             .ColAlignment(8) = flexAlignCenterCenter:
            .TextArray(9) = "색상수":                           .ColAlignment(9) = flexAlignRightCenter:
            .TextArray(10) = "주문량":                          .ColAlignment(10) = flexAlignRightCenter:
            .TextArray(11) = "입고일자":                        .ColAlignment(11) = flexAlignLeftCenter:
            .TextArray(12) = "입고처":                          .ColAlignment(12) = flexAlignRightCenter:
            .TextArray(13) = "입고" & vbCrLf & "절수":          .ColAlignment(13) = flexAlignRightCenter:
            .TextArray(14) = "입고량":                          .ColAlignment(14) = flexAlignRightCenter:
            .TextArray(15) = "배색량":                          .ColAlignment(15) = flexAlignRightCenter:
            
            .TextArray(16) = "Sort OrderID":                    .ColAlignment(16) = flexAlignCenterCenter
            .TextArray(17) = "StuffINSeq":                      .ColAlignment(17) = flexAlignCenterCenter

            .ColWidth(1) = 0
            .ColWidth(2) = 1400
            .ColWidth(3) = 900
            .ColWidth(4) = 1750
            .ColWidth(5) = LIMIT_WIDTH2
            .ColWidth(6) = 500
            .ColWidth(7) = 800
            .ColWidth(8) = 500
            .ColWidth(9) = 600
            .ColWidth(10) = 800
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 650
            .ColWidth(14) = 700
            .ColWidth(15) = 700
            .ColWidth(16) = 0
            .ColWidth(17) = 0
            .ColWidth(18) = 0
            .ColWidth(19) = 0
            .ColWidth(20) = 0
            
            For i = 1 To .Cols - 1
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
            Next i
            .Redraw = flexRDDirect
        End With

    Else
        With grdGroup
            .Redraw = flexRDNone
            
            .FixedRows = 1
            .FixedCols = 0
            .Rows = 1
            .Cols = 21
    
            .TextArray(0) = "":                     .ColWidth(0) = 250
            .TextArray(1) = "":                     .ColWidth(1) = 400
            .TextArray(2) = "거래처ID":             .ColWidth(2) = 0:           .ColAlignment(2) = flexAlignCenterCenter
            .TextArray(3) = "거래처명":             .ColWidth(3) = 1200:        .ColAlignment(3) = flexAlignCenterCenter
            .TextArray(4) = "관리번호":             .ColWidth(4) = 1200:        .ColAlignment(4) = flexAlignCenterCenter
            .TextArray(5) = "Order NO":             .ColWidth(5) = 1200:           .ColAlignment(5) = flexAlignLeftCenter
            .TextArray(6) = "접수일자" & vbCrLf & "입고일자":                   .ColWidth(6) = 900:        .ColAlignment(6) = flexAlignCenterCenter
            .TextArray(7) = "품명" & vbCrLf & "입고처":                         .ColWidth(7) = LIMIT_WIDTH4:        .ColAlignment(7) = flexAlignLeftCenter
            .TextArray(8) = "구분":                 .ColWidth(8) = 630:         .ColAlignment(8) = flexAlignRightCenter
            .TextArray(9) = "원단폭":               .ColWidth(9) = 600:         .ColAlignment(9) = flexAlignCenterCenter
            .TextArray(10) = "축율" & vbCrLf & "LOSS":                .ColWidth(10) = 700:        .ColAlignment(10) = flexAlignCenterCenter
            .TextArray(11) = "색상수":              .ColWidth(11) = 600:        .ColAlignment(11) = flexAlignRightCenter
            .TextArray(12) = "주문량":              .ColWidth(12) = 800:        .ColAlignment(12) = flexAlignRightCenter
            .TextArray(13) = "입고일자":            .ColWidth(13) = 0:          .ColAlignment(13) = flexAlignLeftCenter
            .TextArray(14) = "입고처":              .ColWidth(14) = 0:          .ColAlignment(14) = flexAlignRightCenter
            .TextArray(15) = "입고" & vbCrLf & "절수": .ColWidth(15) = 600:     .ColAlignment(15) = flexAlignRightCenter
            .TextArray(16) = "입고량":              .ColWidth(16) = 860:        .ColAlignment(16) = flexAlignRightCenter
            .TextArray(17) = "배색량":              .ColWidth(17) = 860:        .ColAlignment(17) = flexAlignRightCenter:
            
            
            .TextArray(18) = "거래처ID":            .ColWidth(18) = 0:          .ColAlignment(18) = flexAlignCenterCenter
            .TextArray(19) = "Sort OrderID":        .ColWidth(19) = 0:          .ColAlignment(19) = flexAlignCenterCenter
            .TextArray(20) = "StuffINSeq":          .ColWidth(20) = 0:          .ColAlignment(20) = flexAlignCenterCenter
    
            For i = 1 To .Cols - 1
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
            Next i
    
            .Redraw = flexRDDirect
        End With
    End If
End Sub


'--------------------------------------------------------------------
'  OrderID로 입고와 관련된 Order 데이터 1건 나타내기
'---------------------------------------------------------------------
Private Sub FillStuffOrderData(ByVal OrderID As String)
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim rs As ADODB.Recordset
    Dim lNowRow%, sUnit$
    
    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass

    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    Set rs = oStuffIn.GetStuffINByOrder(OrderID)
    
    Set oStuffIn = Nothing
    
    grdData.Rows = grdData.FixedRows
    
    If rs.RecordCount = 0 Then
        rs.Close
        Set rs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With grdData
        .Redraw = flexRDNone

        lNowRow = IIf(.Rows > .FixedRows, .Row, .FixedRows)
        .Rows = .FixedRows
        sUnit = rs!UnitClss
        
        .AddItem CStr(.Rows) & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                 MakeDate(DF_LONG, rs!AcptDate) & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & _
                 rs!WorkName & vbTab & rs!Width & vbTab & MakeRating(rs!ChunkRate, rs!LossRate) & vbTab & _
                 CheckNum(rs!ColorCnt) & vbTab & SetCurrency(CheckNum(rs!OrderQty)) & vbTab & _
                 SetCurrency(CheckNum(rs!INQty)) & vbTab & rs!UnitClss & vbTab & CheckNum(rs!InRoll)
        
        txtCustomID.Tag = CheckNull(rs!CustomID)
        txtCustomID.Text = CheckNull(rs!kCustom)
        TxtArticleID2.Text = CheckNull(rs!Article)
        TxtArticleID2.Tag = CheckNull(rs!ArticleID)
        txtOrderNO.Text = CheckNull(rs!OrderNo)
        .Redraw = flexRDDirect
    End With
    Call ChangeScroll(0)
    grdData.Editable = flexEDKbdMouse
    rs.Close
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    Call ErrorBox(Err.Number, "frmStuffIN.FillStuffOrderData", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
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
            '.Cell(flexcpFontBold, iRow, 0, iRow, .Cols - 1) = True
        Case 1, 2
            .Cell(flexcpBackColor, iRow, 0, iRow, .Cols - 1) = &HE0E0E0
        End Select
    End With
End Sub



Private Sub grdGroup_DblClick()
    With grdGroup
        If .Row <= .FixedRows Then Exit Sub

        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If
    End With
End Sub

Private Sub grdStuffINSub_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With grdStuffINSub
        If Row < .FixedRows Then Exit Sub
        
        If Col = 11 Then
            If Row = .Rows - 1 Then
                .Rows = .Rows + 1
                .Select .Rows - 1, 1
            End If
        Else
            .Select Row, Col + 1
        End If
    End With
    
    Call ChangeScroll(2)
    Call CalcQty
End Sub
'*******************************************************************************************
'--- 생지 입고 오더별, 거래처별 조회
'---
'
'*******************************************************************************************
Private Sub FillGridGroup(Optional NewValue As Boolean = True)
    Dim oStuffIn As PlusLib2.CStuffIN
    Dim rs As ADODB.Recordset
    Dim iTop(2) As Integer
    Dim i%

  '  On Error GoTo ErrHandler

    Set oStuffIn = New PlusLib2.CStuffIN
    oStuffIn.Connection = g_adoCon
    oStuffIn.UserName = g_sUserName
    
    Set rs = oStuffIn.GetStuffIN(IIf(chkSearch(3) = vbChecked, 1, 0) _
                                , MakeDate(DF_SHORT, dtpDate(0)) _
                                , MakeDate(DF_SHORT, dtpDate(1)) _
                                , IIf(chkSearch(1) = vbChecked, 1, 0) _
                                , txtCustom(1).Tag _
                                , IIf(chkSearch(2) = vbChecked, 1, 0) _
                                , txtArticle.Tag _
                                , IIf(chkSearch(4) = vbChecked, 1, 0) _
                                , Left(CboStuffClss2, 1))

    Set oStuffIn = Nothing
    
    If rs.RecordCount = 0 Then
        grdGroup.Rows = grdGroup.FixedRows
        Exit Sub
    End If
    
    Call InitGroup(NewValue)
    
    If NewValue Then
        
        With grdGroup
            .Redraw = flexRDNone
            .Rows = .FixedRows

            Do Until rs.EOF
                
                If rs!OrderID <> .TextMatrix(.Rows - 1, 16) Then
                    .AddItem " " & vbTab & _
                                MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & _
                                IIf(Trim(rs!OrderNo) = "", "", rs!OrderNo) & vbTab & _
                                MakeDate(DF_MID, rs!AcptDate) & vbTab & _
                                rs!Custom1 & vbTab & _
                                rs!Article & vbTab & _
                                rs!WorkName & vbTab & _
                                rs!Width & vbTab & _
                                MakeRating(rs!ChunkRate, rs!LossRate) & vbTab & _
                                CheckNum(rs!ColorQty) & vbTab & _
                                SetCurrency(CheckNum(rs!OrderQty)) & vbTab & _
                                "" & vbTab & _
                                "" & vbTab & _
                                "0" & vbTab & _
                                "0" & vbTab & _
                                rs!배색Qty & vbTab & _
                                rs!OrderID & vbTab & _
                                rs!StuffDate & vbTab & _
                                rs!OrderNo & vbTab & _
                                "" & vbTab & ""
                
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 1)
                    iTop(1) = .Rows - 1
                
                End If
'
                .AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                         CheckNull(rs!Custom2) & vbTab & _
                         MakeDate(DF_MID, rs!StuffDate) & "(" + CheckNull(rs!ThreadName) + ")" & vbTab & _
                         "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                         rs!StuffRoll & vbTab & SetCurrency(rs!StuffQty) & vbTab & "0" & vbTab & rs!OrderID & vbTab & _
                         rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                         
                ' OrderId와 입고일자 + 구분 + 일련번호를 설정한다.
                .TextMatrix(.Rows - 1, .Cols - 2) = rs!OrderID
                .TextMatrix(.Rows - 1, .Cols - 1) = rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                
                '-------입고절수 , 입고수량 Order별로 합계
                .TextMatrix(iTop(1), 13) = SetCurrency(.TextMatrix(iTop(1), 13) + rs!StuffRoll)
                .TextMatrix(iTop(1), 14) = SetCurrency(.TextMatrix(iTop(1), 14) + rs!StuffQty)
    
                rs.MoveNext
            Loop
            
            Call ChangeScroll(1)
            
            .Redraw = flexRDDirect
        End With
        
    Else
        
        With grdGroup
            .Redraw = flexRDNone
            .Rows = 1
            
            Do Until rs.EOF
                
                If Trim(rs!CustomID1) <> Trim(.TextMatrix(.Rows - 1, 18)) Then
                    .AddItem "" & vbTab & "" & vbTab & rs!CustomID1 & vbTab & rs!Custom1 & vbTab & "" & vbTab & _
                    "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                    "" & vbTab & "" & vbTab & "" & vbTab & _
                    "" & vbTab & "" & vbTab & "" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & _
                    rs!CustomID1 & vbTab & "" & vbTab & rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 1)
                    
                    iTop(1) = .Rows - 1
                
                End If
                
                If Trim(rs!OrderID) <> Trim(.TextMatrix(.Rows - 1, 19)) Then
                    .AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!OrderNo & vbTab & _
                    MakeDate(DF_MID, rs!AcptDate) & vbTab & rs!Article & vbTab & rs!WorkName & vbTab & _
                    rs!Width & vbTab & MakeRating(rs!ChunkRate, rs!LossRate) & vbTab & CheckNum(rs!ColorQty) & vbTab & SetCurrency(CheckNum(rs!OrderQty)) & vbTab & _
                    "" & vbTab & "" & vbTab & "0" & vbTab & "0" & vbTab & rs!배색Qty & vbTab & _
                    rs!CustomID1 & vbTab & rs!OrderID & vbTab & rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                
                    Call DoFlexGridGroup(grdGroup, .Rows - 1, 2)
                    
                    iTop(2) = .Rows - 1
                
                End If
                
                
                .AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                         MakeDate(DF_MID, rs!StuffDate) & vbTab & CheckNull(rs!Custom2) & "(" & rs!ThreadName & ")" & vbTab & "" & vbTab & _
                           "" & vbTab & "" & vbTab & "" & vbTab & _
                           "" & vbTab & "" & vbTab & "" & vbTab & _
                        rs!StuffRoll & vbTab & SetCurrency(rs!StuffQty) & vbTab & "0" & vbTab & _
                        rs!CustomID1 & vbTab & rs!OrderID & vbTab & rs!StuffDate + "-" + rs!StuffClss + "-" + CStr(rs!StuffSeq)
                
                
                For i = 1 To 2
                    .TextMatrix(iTop(i), 15) = SetCurrency(.TextMatrix(iTop(i), 15) + rs!StuffRoll)
                    .TextMatrix(iTop(i), 16) = SetCurrency(.TextMatrix(iTop(i), 16) + rs!StuffQty)
                Next i

                rs.MoveNext
            Loop
            
            Call ChangeScroll(1)
            
            .Redraw = flexRDDirect
        End With
    End If
    
    If grdGroup.Rows > grdGroup.FixedRows Then
        grdGroup.Row = grdGroup.FixedRows
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub

ErrHandler:
    grdGroup.Redraw = flexRDDirect
    Call ErrorBox(Err.Number, "frmStuffIN.FillGridGroup", Err.Description)
    Set rs = Nothing
    Set oStuffIn = Nothing
End Sub

Private Sub grdStuffINSub_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    '---- 1번 컬럼은 Validateedit에서 확인 하지 않음.
    
    With grdStuffINSub
        If Col <> 1 Then
            If Not IsNumeric(.EditText) Then
                Cancel = True
            End If

            If .EditText = "0" Then
                Cancel = True
            End If

            If InStr(.EditText, "*") > 0 Then
                Cancel = False
            End If

            If Len(.EditText) = 0 Then
                Cancel = False
            End If

        End If
    End With
    
End Sub

Private Sub optGroup_Click(Index As Integer)
    If Index = 0 Then
        m_bGroupClss = True
        Call InitGroup
    Else
        m_bGroupClss = False
        Call InitGroup(m_bGroupClss)
    End If
    
    If tabForm.Tab = 0 Then
        FillGridGroup (m_bGroupClss)
    End If
    Call optOrder_Click(Index)
End Sub

Private Sub optOrder_Click(Index As Integer)
    Select Case Index
''        Case 0
''            chkSearch(0).Caption = optOrder(Index).Caption
''            grdData.ColWidth(1) = 0
''            grdData.ColWidth(2) = 1300
''        Case 1
''            pnlCaption(4) = "관리 번호"
''            chkSearch(0).Caption = optOrder(Index).Caption
''
''            grdData.ColWidth(1) = 1300
''            grdData.ColWidth(2) = 0
        Case 2
            With grdGroup
                If m_bGroupClss Then
                    .TextMatrix(0, 1) = ""
                    .ColWidth(1) = 200
                    .ColWidth(2) = 1300
                Else
                    .ColWidth(4) = 0
                    .ColWidth(5) = 1300
                End If
            End With
        Case 3
            With grdGroup
                If m_bGroupClss Then
                    .TextMatrix(0, 1) = "관리 번호"
                    .ColWidth(1) = 1300
                    .ColWidth(2) = 0
                Else
                    .ColWidth(4) = 1300
                    .ColWidth(5) = 0
                End If
            End With
    End Select
End Sub

Private Sub tabform_Click(PreviousTab As Integer)
    Select Case tabForm.Tab
        Case 0
            Call SetClearEdit
            Call cmdSearch_Click
        Case 1
            If Trim(pnlMsg) = "" Or val(txtStuffSeq.Text) = 0 Then
                Call cmdOperate_Click(0)
            End If
    End Select

End Sub

Private Sub txtArticle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_ARTICLE, , False, txtArticle)
        Call MoveFocus(KeyAscii)
    End If

End Sub




Private Sub TxtArticleID2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_ARTICLE, , False, TxtArticleID2)
        Call MoveFocus(KeyAscii)
    End If
End Sub

Private Sub txtCustom_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            Call MoveFocus(KeyAscii)
        Case 1
            If KeyAscii = vbKeyReturn Then
                Call ReturnCode(LG_CUSTOM, , False, txtCustom(Index))
                Call MoveFocus(KeyAscii)
            End If
    End Select
End Sub

Private Sub txtCustomID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ReturnCode(LG_CUSTOM, , False, txtCustomID)
        Call MoveFocus(KeyAscii)
    End If
End Sub



Private Sub txtNum_Change(Index As Integer)
    txtNum(Index) = SetCurrency(txtNum(Index))
End Sub

Private Sub txtNum_GotFocus(Index As Integer)
    Call GotFocusText(txtNum(Index))
End Sub

Private Sub txtNum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call NextFocus
    End If
End Sub

Private Sub txtNum_LostFocus(Index As Integer)
    txtNum(Index) = SetCurrency(txtNum(Index))
End Sub


Private Sub txtOrderNO_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)
End Sub

Private Sub txtOrderNO_LostFocus()
    If Len(txtOrderNO) > 0 Then
        Call cmdFind_Click(3)
    End If
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call FillStuffOrderData(txtSearch)
        KeyAscii = 0
'        Call cmdFind_Click(3)
        cmdSearch.SetFocus
    End If
End Sub

Private Sub txtSearch_LostFocus()
    If Len(txtSearch) = 10 Then
        Call FillStuffOrderData(txtSearch)
    Else
        txtSearch = ""
        grdData.Rows = grdData.FixedRows
    End If

End Sub

Private Sub txtThreadName_KeyPress(KeyAscii As Integer)
    Call MoveFocus(KeyAscii)
End Sub
