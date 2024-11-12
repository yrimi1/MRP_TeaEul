VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSetting 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "자사정보(0006)"
   ClientHeight    =   8805
   ClientLeft      =   2055
   ClientTop       =   3105
   ClientWidth     =   10680
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7875
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   3555
      _cx             =   6271
      _cy             =   13891
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSetting.frx":000C
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
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   9210
      TabIndex        =   32
      Top             =   8070
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlEdit 
      Height          =   7920
      Left            =   3660
      TabIndex        =   14
      Top             =   90
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   13970
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Frame fraMoreInfo 
         Caption         =   "추가정보"
         Height          =   8025
         Left            =   6960
         TabIndex        =   95
         Top             =   90
         Visible         =   0   'False
         Width           =   4755
         Begin VB.Frame fraSMS2 
            Caption         =   "SMS서버2"
            Height          =   1605
            Left            =   120
            TabIndex        =   120
            Top             =   6360
            Width           =   4575
            Begin VB.TextBox txtSMS2Data 
               Height          =   300
               Index           =   0
               Left            =   975
               TabIndex        =   64
               Top             =   240
               Width           =   3510
            End
            Begin VB.TextBox txtSMS2Data 
               Height          =   300
               Index           =   1
               Left            =   975
               TabIndex        =   65
               Top             =   570
               Width           =   1230
            End
            Begin VB.TextBox txtSMS2Data 
               Height          =   300
               Index           =   2
               Left            =   3240
               TabIndex        =   66
               Top             =   570
               Width           =   1230
            End
            Begin VB.TextBox txtSMS2Data 
               Height          =   300
               Index           =   3
               Left            =   975
               TabIndex        =   67
               Top             =   900
               Width           =   1230
            End
            Begin VB.TextBox txtSMS2Data 
               Height          =   300
               Index           =   4
               Left            =   3255
               TabIndex        =   68
               Top             =   900
               Width           =   1230
            End
            Begin VB.TextBox txtSMS2Data 
               Height          =   300
               Index           =   5
               Left            =   975
               TabIndex        =   69
               Top             =   1230
               Width           =   3510
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   43
               Left            =   90
               TabIndex        =   121
               Top             =   570
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "포트From"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   44
               Left            =   2370
               TabIndex        =   122
               Top             =   570
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "포트To"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   45
               Left            =   90
               TabIndex        =   123
               Top             =   900
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "로그인ID"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   46
               Left            =   2370
               TabIndex        =   124
               Top             =   900
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "암호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   47
               Left            =   90
               TabIndex        =   125
               Top             =   1230
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "인증코드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   48
               Left            =   90
               TabIndex        =   126
               Top             =   240
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "주소"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin VB.Frame fraSMS1 
            Caption         =   "SMS서버1"
            Height          =   1665
            Left            =   120
            TabIndex        =   114
            Top             =   4680
            Width           =   4575
            Begin VB.TextBox txtSMS1Data 
               Height          =   300
               Index           =   0
               Left            =   975
               TabIndex        =   58
               Top             =   270
               Width           =   3510
            End
            Begin VB.TextBox txtSMS1Data 
               Height          =   300
               Index           =   5
               Left            =   975
               TabIndex        =   63
               Top             =   1260
               Width           =   3510
            End
            Begin VB.TextBox txtSMS1Data 
               Height          =   300
               Index           =   4
               Left            =   3255
               TabIndex        =   62
               Top             =   930
               Width           =   1230
            End
            Begin VB.TextBox txtSMS1Data 
               Height          =   300
               Index           =   3
               Left            =   975
               TabIndex        =   61
               Top             =   930
               Width           =   1230
            End
            Begin VB.TextBox txtSMS1Data 
               Height          =   300
               Index           =   2
               Left            =   3240
               TabIndex        =   60
               Top             =   600
               Width           =   1230
            End
            Begin VB.TextBox txtSMS1Data 
               Height          =   300
               Index           =   1
               Left            =   975
               TabIndex        =   59
               Top             =   600
               Width           =   1230
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   38
               Left            =   90
               TabIndex        =   115
               Top             =   600
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "포트From"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   39
               Left            =   2370
               TabIndex        =   116
               Top             =   600
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "포트To"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   40
               Left            =   90
               TabIndex        =   117
               Top             =   930
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "로그인ID"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   41
               Left            =   2370
               TabIndex        =   118
               Top             =   930
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "암호"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   42
               Left            =   90
               TabIndex        =   119
               Top             =   1260
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "인증코드"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   49
               Left            =   90
               TabIndex        =   127
               Top             =   270
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "주소"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin VB.Frame fraFTP 
            Caption         =   "FTP"
            Height          =   2355
            Left            =   90
            TabIndex        =   105
            Top             =   2310
            Width           =   4575
            Begin VB.TextBox txtFTPData 
               Height          =   300
               Index           =   0
               Left            =   975
               TabIndex        =   49
               Top             =   240
               Width           =   3510
            End
            Begin VB.TextBox txtFTPData 
               Height          =   300
               Index           =   1
               Left            =   975
               TabIndex        =   50
               Top             =   570
               Width           =   1230
            End
            Begin VB.TextBox txtFTPData 
               Height          =   300
               Index           =   2
               Left            =   3240
               TabIndex        =   51
               Top             =   570
               Width           =   1230
            End
            Begin VB.TextBox txtFTPData 
               Height          =   300
               Index           =   3
               Left            =   975
               TabIndex        =   52
               Top             =   900
               Width           =   1230
            End
            Begin VB.TextBox txtFTPData 
               Height          =   300
               Index           =   4
               Left            =   3255
               TabIndex        =   53
               Top             =   900
               Width           =   1230
            End
            Begin VB.TextBox txtFTPData 
               Height          =   300
               Index           =   5
               Left            =   975
               TabIndex        =   54
               Top             =   1230
               Width           =   3510
            End
            Begin VB.TextBox txtFTPData 
               Height          =   300
               Index           =   6
               Left            =   975
               TabIndex        =   55
               Top             =   1650
               Width           =   1230
            End
            Begin VB.TextBox txtFTPData 
               Height          =   300
               Index           =   7
               Left            =   3255
               TabIndex        =   56
               Top             =   1620
               Width           =   1230
            End
            Begin VB.TextBox txtFTPData 
               Height          =   300
               Index           =   8
               Left            =   975
               TabIndex        =   57
               Top             =   1980
               Width           =   3510
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   30
               Left            =   90
               TabIndex        =   106
               Top             =   570
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "포트From"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   31
               Left            =   2370
               TabIndex        =   107
               Top             =   570
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "포트To"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   32
               Left            =   90
               TabIndex        =   108
               Top             =   900
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "로그인ID1"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   33
               Left            =   2370
               TabIndex        =   109
               Top             =   900
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "암호1"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   34
               Left            =   90
               TabIndex        =   110
               Top             =   1230
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "인증코드1"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   35
               Left            =   90
               TabIndex        =   111
               Top             =   1620
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "로그인ID2"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   36
               Left            =   2370
               TabIndex        =   112
               Top             =   1620
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "암호2"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   37
               Left            =   90
               TabIndex        =   113
               Top             =   1950
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "인증코드2"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   50
               Left            =   90
               TabIndex        =   128
               Top             =   240
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "주소"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
         Begin VB.Frame fraWebInfo1 
            Caption         =   "Web"
            Height          =   2025
            Left            =   90
            TabIndex        =   96
            Top             =   270
            Width           =   4575
            Begin VB.TextBox txtWebData 
               Height          =   300
               Index           =   7
               Left            =   975
               TabIndex        =   48
               Top             =   1620
               Width           =   3510
            End
            Begin VB.TextBox txtWebData 
               Height          =   300
               Index           =   6
               Left            =   3255
               TabIndex        =   47
               Top             =   1290
               Width           =   1230
            End
            Begin VB.TextBox txtWebData 
               Height          =   300
               Index           =   5
               Left            =   975
               TabIndex        =   46
               Top             =   1320
               Width           =   1230
            End
            Begin VB.TextBox txtWebData 
               Height          =   300
               Index           =   4
               Left            =   975
               TabIndex        =   45
               Top             =   900
               Width           =   3510
            End
            Begin VB.TextBox txtWebData 
               Height          =   300
               Index           =   3
               Left            =   3255
               TabIndex        =   44
               Top             =   570
               Width           =   1230
            End
            Begin VB.TextBox txtWebData 
               Height          =   300
               Index           =   2
               Left            =   975
               TabIndex        =   43
               Top             =   570
               Width           =   1230
            End
            Begin VB.TextBox txtWebData 
               Height          =   300
               Index           =   1
               Left            =   3240
               TabIndex        =   42
               Top             =   240
               Width           =   1230
            End
            Begin VB.TextBox txtWebData 
               Height          =   300
               Index           =   0
               Left            =   975
               TabIndex        =   41
               Top             =   240
               Width           =   1230
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   22
               Left            =   90
               TabIndex        =   97
               Top             =   240
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "포트From"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   23
               Left            =   2370
               TabIndex        =   98
               Top             =   240
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "포트To"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   24
               Left            =   90
               TabIndex        =   99
               Top             =   570
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "로그인ID1"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   25
               Left            =   2370
               TabIndex        =   100
               Top             =   570
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "암호1"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   26
               Left            =   90
               TabIndex        =   101
               Top             =   900
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "인증코드1"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   27
               Left            =   90
               TabIndex        =   102
               Top             =   1290
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "로그인ID2"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   28
               Left            =   2370
               TabIndex        =   103
               Top             =   1290
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "암호2"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   29
               Left            =   90
               TabIndex        =   104
               Top             =   1620
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "인증코드2"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
         End
      End
      Begin VB.Frame fraBaseInfo 
         Caption         =   "기본정보"
         Height          =   8265
         Left            =   60
         TabIndex        =   37
         Top             =   30
         Width           =   6885
         Begin VB.Frame fraAddress 
            Height          =   2235
            Left            =   60
            TabIndex        =   130
            Top             =   3210
            Width           =   6765
            Begin VB.Frame fraDoro 
               Caption         =   "도로명주소"
               Height          =   1155
               Left            =   60
               TabIndex        =   132
               Top             =   150
               Width           =   6645
               Begin VB.TextBox txtAddressAssist 
                  Height          =   300
                  Left            =   930
                  TabIndex        =   17
                  Top             =   780
                  Width           =   5670
               End
               Begin VB.TextBox txtAddress1 
                  BackColor       =   &H00C0FFC0&
                  Height          =   300
                  Left            =   930
                  TabIndex        =   15
                  Top             =   180
                  Width           =   5670
               End
               Begin VB.TextBox txtAddress2 
                  Height          =   300
                  Left            =   930
                  TabIndex        =   16
                  Top             =   480
                  Width           =   5670
               End
               Begin VB.TextBox txtGunMoolMngNo 
                  Enabled         =   0   'False
                  Height          =   270
                  Left            =   120
                  TabIndex        =   133
                  TabStop         =   0   'False
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin Threed.SSPanel pnlCaption 
                  Height          =   300
                  Index           =   4
                  Left            =   60
                  TabIndex        =   134
                  Top             =   780
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  _Version        =   196609
                  Caption         =   "보조주소"
                  BevelOuter      =   1
                  RoundedCorners  =   0   'False
                  FloodShowPct    =   -1  'True
               End
            End
            Begin VB.Frame fraJiBun 
               Caption         =   "지번주소"
               Height          =   825
               Left            =   60
               TabIndex        =   131
               Top             =   1320
               Width           =   6645
               Begin VB.TextBox txtName 
                  BackColor       =   &H00C0FFC0&
                  Height          =   300
                  Index           =   7
                  Left            =   990
                  TabIndex        =   18
                  Top             =   120
                  Width           =   5610
               End
               Begin VB.TextBox txtName 
                  Height          =   300
                  Index           =   8
                  Left            =   990
                  TabIndex        =   19
                  Top             =   450
                  Width           =   5610
               End
            End
         End
         Begin VB.Frame fraOldNNew 
            Height          =   405
            Left            =   2760
            TabIndex        =   129
            Top             =   2820
            Width           =   1875
            Begin VB.OptionButton optOldNNew 
               Caption         =   "지번"
               Height          =   225
               Index           =   1
               Left            =   1050
               TabIndex        =   13
               Top             =   150
               Width           =   675
            End
            Begin VB.OptionButton optOldNNew 
               Caption         =   "도로명"
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   12
               Top             =   150
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1125
            TabIndex        =   1
            Top             =   210
            Width           =   1230
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   2
            Left            =   1125
            TabIndex        =   3
            Top             =   870
            Width           =   5715
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   3
            Left            =   1125
            TabIndex        =   4
            Top             =   1200
            Width           =   2370
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   4
            Left            =   4605
            TabIndex        =   5
            Top             =   1200
            Width           =   2220
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   5
            Left            =   1125
            TabIndex        =   8
            Top             =   2190
            Width           =   3975
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   6
            Left            =   1125
            TabIndex        =   9
            Top             =   2535
            Width           =   3975
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Index           =   1
            Left            =   1125
            TabIndex        =   2
            Top             =   540
            Width           =   5730
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   14
            Left            =   1125
            TabIndex        =   24
            Top             =   6870
            Width           =   4710
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   13
            Left            =   1125
            TabIndex        =   25
            Top             =   6510
            Width           =   4710
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   12
            Left            =   1125
            TabIndex        =   23
            Top             =   6150
            Width           =   4710
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   11
            Left            =   1125
            TabIndex        =   22
            Top             =   5820
            Width           =   2400
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   10
            Left            =   4605
            TabIndex        =   21
            Top             =   5490
            Width           =   2190
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   9
            Left            =   1125
            TabIndex        =   20
            Top             =   5490
            Width           =   2400
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   15
            Left            =   1125
            TabIndex        =   28
            Top             =   7200
            Width           =   4710
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   16
            Left            =   1125
            TabIndex        =   29
            Top             =   7500
            Width           =   4710
         End
         Begin VB.Frame fraUseYN 
            Caption         =   "사용여부"
            Enabled         =   0   'False
            Height          =   915
            Left            =   5850
            TabIndex        =   40
            Top             =   6930
            Width           =   1005
            Begin VB.OptionButton optUseYn 
               Caption         =   "예"
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   30
               Top             =   240
               Value           =   -1  'True
               Width           =   885
            End
            Begin VB.OptionButton optUseYn 
               Caption         =   "아니오"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   31
               Top             =   540
               Width           =   885
            End
         End
         Begin VB.Frame fraRPYN 
            Caption         =   "대표사용"
            Enabled         =   0   'False
            Height          =   945
            Left            =   5850
            TabIndex        =   39
            Top             =   5940
            Width           =   1005
            Begin VB.OptionButton optRPYn 
               Caption         =   "아니오"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   27
               Top             =   540
               Width           =   885
            End
            Begin VB.OptionButton optRPYn 
               Caption         =   "예"
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   26
               Top             =   240
               Value           =   -1  'True
               Width           =   885
            End
         End
         Begin VB.TextBox txtRPYN_OLD 
            Height          =   345
            Left            =   4770
            TabIndex        =   38
            Top             =   5820
            Visible         =   0   'False
            Width           =   375
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   855
            Left            =   30
            TabIndex        =   70
            Top             =   7890
            Visible         =   0   'False
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   1508
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txtName 
               Height          =   300
               Index           =   20
               Left            =   1110
               TabIndex        =   72
               Top             =   270
               Width           =   2370
            End
            Begin VB.TextBox txtName 
               Height          =   300
               Index           =   19
               Left            =   1110
               TabIndex        =   71
               Top             =   0
               Width           =   2370
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   11
               Left            =   30
               TabIndex        =   73
               Top             =   0
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "SERVER"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   12
               Left            =   60
               TabIndex        =   74
               Top             =   240
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "DATABASE"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin MSComDlg.CommonDialog dlgLogo 
               Left            =   3000
               Top             =   330
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin Threed.SSPanel pnlCaption 
               Height          =   300
               Index           =   6
               Left            =   0
               TabIndex        =   75
               Top             =   480
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   529
               _Version        =   196609
               Caption         =   "회사 로고"
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSCommand cmdFind 
               Height          =   315
               Index           =   0
               Left            =   1230
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   480
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               _Version        =   196609
               ButtonStyle     =   3
               Outline         =   0   'False
            End
            Begin VB.Image imgLogo 
               BorderStyle     =   1  '단일 고정
               Height          =   1200
               Left            =   3420
               Stretch         =   -1  'True
               Top             =   -450
               Width           =   1230
            End
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   77
            Top             =   540
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "상호(한글)"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   78
            Top             =   1200
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "약어"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSMask.MaskEdBox mskName 
            Height          =   300
            Index           =   0
            Left            =   1125
            TabIndex        =   6
            Top             =   1530
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   12648384
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   5
            Left            =   90
            TabIndex        =   79
            Top             =   1530
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "사업자번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   90
            TabIndex        =   80
            Top             =   2190
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "업  태"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   3
            Left            =   90
            TabIndex        =   81
            Top             =   2550
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "종  목"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   315
            Index           =   1
            Left            =   2430
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2880
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   7
            Left            =   90
            TabIndex        =   82
            Top             =   2895
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "우편번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   8
            Left            =   90
            TabIndex        =   83
            Top             =   5490
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "대표전화"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   9
            Left            =   90
            TabIndex        =   84
            Top             =   5820
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "팩스번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSMask.MaskEdBox mskName 
            Height          =   300
            Index           =   1
            Left            =   1125
            TabIndex        =   10
            Top             =   2895
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            Mask            =   "###-###"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   10
            Left            =   90
            TabIndex        =   85
            Top             =   870
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "상호(영문)"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   13
            Left            =   3570
            TabIndex        =   86
            Top             =   5490
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "전화번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   14
            Left            =   90
            TabIndex        =   87
            Top             =   6855
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "계좌번호1"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   15
            Left            =   90
            TabIndex        =   88
            Top             =   7185
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "계좌번호2"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   16
            Left            =   90
            TabIndex        =   89
            Top             =   7515
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "계좌번호3"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   17
            Left            =   3570
            TabIndex        =   90
            Top             =   1200
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "대 표 자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   18
            Left            =   90
            TabIndex        =   91
            Top             =   210
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "코     드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   19
            Left            =   90
            TabIndex        =   92
            Top             =   6165
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "이메일"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   20
            Left            =   90
            TabIndex        =   93
            Top             =   6525
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "홈페이지"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin MSMask.MaskEdBox mskName 
            Height          =   300
            Index           =   2
            Left            =   1725
            TabIndex        =   7
            Top             =   1860
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "######-#######"
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   21
            Left            =   90
            TabIndex        =   94
            Top             =   1860
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "주민/법인등록번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   690
      Left            =   6270
      TabIndex        =   33
      Top             =   8070
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      저장(&S)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   510
      Left            =   4470
      TabIndex        =   34
      Top             =   8220
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   900
      _Version        =   196609
      BackColor       =   65535
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   690
      Left            =   7740
      TabIndex        =   35
      Top             =   8070
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      취소(&C)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Label lblTip 
      Caption         =   "수정은 그리드 더블클릭!!!"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   36
      Top             =   8250
      Width           =   2535
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************************
'** System 명 : Mrpplus2
'** 모듈명    :
'** Author    : Wizard
'** 작성자    :
'** 내용      :
'** 생성일자  :
'------------------------------------------------------------------------------------------------------------------
' * 변경이력
'------------------------------------------------------------------------------------------------------------------
' 일자        작업자  요청자          요청번호           요청내용 및 변경내용
'------------------------------------------------------------------------------------------------------------------
' 2013.12.12  오승욱                 S_201312_태을염직_99    지번주소에서 도로명 주소로 입력가능하게
'*******************************************************************************

Private m_sPath As String


Private m_sFlag        As String * 1

Private Sub cmdCancel_Click()

    Call ChangeMode(Me, True)
    
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    m_sFlag = ""
    
    Call FillGrid
        
    grdData.SetFocus
End Sub

Private Sub cmdExit_Click()
'S_201110_대진텍스_03 에 의한 수정-주석 처리
'''    Call SaveRegistry  ' 레지스트리에 저장. 쓰지 않음.
''
''   ' If (QuestionBox("변경된 내용을 저장하시겠습니까?")) Then
''   If (MsgBox("변경된 내용을 저장하시겠습니까?", vbYesNo + vbQuestion, "자료 저장") = vbYes) Then
''
''        Call SaveInfo 'db에 저장하는 부분...
''    End If
''
''
''    If Len(Trim(txtName(0))) > 0 Then
''        PlusMDI.Caption = LoadResString(101) & " - " & txtName(0)
''    Else
''        PlusMDI.Caption = LoadResString(101)
''    End If
''
''
    Unload Me

End Sub

'S_201110_대진텍스_03 에 의한 추가
Private Sub ClearData()

    Call ClearText(txtName)
    Call ClearText(mskName)
    
    'S_201312_태을염직_99 에 의한 추가---------------------------------------------
    optOldNNew(0).Value = True
    
    txtGunMoolMngNo.Text = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtAddressAssist.Text = ""
    '-----------------------------------------------------------------------------------
    
End Sub


Private Sub SaveInfo()
    Dim oInfo As PlusLib2.CInfo
    Dim rs As ADODB.Recordset
    Dim oinfotable As TCompanyInfo2
    Dim sFileName$
    
    On Error GoTo ErrSaveInfo:
    
    Set oInfo = New PlusLib2.CInfo

    
    'S_201110_대진텍스_03 에 의한 수정-NEW
    With oinfotable
        .sCompanyID = CheckNull(txtName(0))         ' 자사코드
        .sKCompany = CheckNull(txtName(1))          ' 한글 상호
        .sECompany = CheckNull(txtName(2))          ' 영문 상호
        .sShortCompany = CheckNull(txtName(3))      ' 약어
        .sChief = CheckNull(txtName(4))             ' 대표자
        .sCompanyNo = CheckNull(mskName(0))         ' 사업자번호
        .sRegistID = CheckNull(mskName(2))          '주민등록번호/법인등록번호
        .sCondition = CheckNull(txtName(5))         ' 업태
        .sCategory = CheckNull(txtName(6))          ' 업종
        .sZipCode = CheckNull(mskName(1))           ' 우편번호
                
        'S_201312_태을염직_99 에 의한 추가---------------------------------------------------------
        .sOldNNewClss = IIf(optOldNNew(0).Value = True, "0", "1")                       '도로명,지번주소 구분 0:도로명, 1:지번
        .sGunMoolMngNo = IIf(optOldNNew(0).Value = True, txtGunMoolMngNo.Text, "")        '건물관리 고유식별번호
        .sAddress1 = CheckNull(txtAddress1.Text)                              ' 도로명주소1
        .sAddress2 = CheckNull(txtAddress2.Text)                      ' 도로명주소2
        .sAddressAssist = CheckNull(txtAddressAssist.Text)                  ' 도로명 보조주소
        '-------------------------------------------------------------------------------------
        'S_201312_태을염직_99 에 의한 수정(OLD:sAddress1)
        .sAddressJiBun1 = CheckNull(txtName(7))          ' 지번주소1
        'S_201312_태을염직_99 에 의한 수정(OLD:sAddress2)
        .sAddressJiBun2 = CheckNull(txtName(8))          ' 지번주소2

        .sPhone1 = CheckNull(txtName(9))            ' 대표전화
        .sPhone2 = CheckNull(txtName(10))           ' 전화번호
        .sFaxNO = CheckNull(txtName(11))            ' 팩스번호
        .sEMail = CheckNull(txtName(12))            ' 이메일
        .sHomePage = CheckNull(txtName(13))         ' 홈페이지
        
        '추가정보*********************************************************
        '2012.03.19 추가
        ' --WebPage로그인정보
        .sWebPortFrom = CheckNull(txtWebData(0).Text)       ' WebPage포트From
        .sWebPortTo = CheckNull(txtWebData(1).Text)         ' WebPage포트To
        .sWebID1 = CheckNull(txtWebData(2).Text)            ' WebPa ge로그인ID1
        .sWebPass1 = CheckNull(txtWebData(3).Text)          ' WebPage로그인암호1
        .sWebAuthCode1 = CheckNull(txtWebData(4).Text)      ' WebPage로그인인증코드1
        .sWebID2 = CheckNull(txtWebData(5).Text)            ' WebPage로그인ID2
        .sWebPass2 = CheckNull(txtWebData(6).Text)          ' WebPage로그인암호2
        .sWebAuthCode2 = CheckNull(txtWebData(7).Text)      ' WebPage로그인인증코드2
        
        ' --FTP로그인정보
        .sFTPPage = CheckNull(txtFTPData(0).Text)           ' FTP주소
        .sFTPPortFrom = CheckNull(txtFTPData(1).Text)       ' FTP포트From
        .sFTPPortTo = CheckNull(txtFTPData(2).Text)         ' FTP포트To
        .sFTPID1 = CheckNull(txtFTPData(3).Text)            ' FTP로그인ID1
        .sFTPPass1 = CheckNull(txtFTPData(4).Text)          ' FTP로그인암호1
        .sFTPAuthCode1 = CheckNull(txtFTPData(5).Text)      ' FTP로그인인증코드1
        .sFTPID2 = CheckNull(txtFTPData(6).Text)            ' FTP로그인ID2
        .sFTPPass2 = CheckNull(txtFTPData(7).Text)          ' FTP로그인암호2
        .sFTPAuthCode2 = CheckNull(txtFTPData(8).Text)      ' FTP로그인인증코드2
        
        ' --SMS서버1그인정보
        .sSMSURL1 = CheckNull(txtSMS1Data(0).Text)          ' 문자전송서버1주소
        .sSMSPortFrom1 = CheckNull(txtSMS1Data(1).Text)     ' 문자전송서버1포트From
        .sSMSPortTo1 = CheckNull(txtSMS1Data(2).Text)       ' 문자전송서버1포트To
        .sSMSID1 = CheckNull(txtSMS1Data(3).Text)           ' 문자전송서버1아이디
        .sSMSPASS1 = CheckNull(txtSMS1Data(4).Text)         ' 문자전송서버1암호
        .sSMSAuthCode1 = CheckNull(txtSMS1Data(5).Text)     ' 문자전송서버1인증코드
        
        ' --SMS서버1그인정보
        .sSMSURL2 = CheckNull(txtSMS2Data(0).Text)          ' 문자전송서버2주소
        .sSMSPortFrom2 = CheckNull(txtSMS2Data(1).Text)     ' 문자전송서버2포트From
        .sSMSPortTo2 = CheckNull(txtSMS2Data(2).Text)       ' 문자전송서버2포트To
        .sSMSID2 = CheckNull(txtSMS2Data(3).Text)           ' 문자전송서버2아이디
        .sSMSPASS2 = CheckNull(txtSMS2Data(4).Text)         ' 문자전송서버2암호
        .sSMSAuthCode2 = CheckNull(txtSMS2Data(5).Text)     ' 문자전송서버2인증코드
       '*****************************************************************

        .sBank1 = CheckNull(txtName(14))            ' 계좌번호1
        .sBank2 = CheckNull(txtName(15))            ' 계좌번호2
        .sBank3 = CheckNull(txtName(16))            ' 계좌번호3
        .sRPYn = IIf(optRPYn(0).Value = True, "Y", "N") '대표여부
        .sUseYn = IIf(optUseYn(0).Value = True, "Y", "N") '사용여부
        
        .sRPYn_OLD = txtRPYN_OLD.Text       '이전 대표여부 설정값
    End With
    
    
    sFileName = GetWindowsPath & "\Wizard.ini"
    
    oInfo.Connection = g_adoCon
    oInfo.UserName = g_sUserName
    
    If oInfo.SaveCompanyInfo(oinfotable) Then    ' 자료 저장
''        If imgLogo.Picture <> "0" Then Call oInfo.SaveCompanyLogo(m_sPath)
        MsgBox "정상적으로 변경 되었습니다"

        Call ChangeMode(Me, True)
        
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
    
        m_sFlag = ""
        grdData.SetFocus
    
        Exit Sub
        
    End If

ErrSaveInfo:
    Set rs = Nothing
    Set oInfo = Nothing
    Call ErrorBox(Err.Number, "Setting.SaveInfo", Err.Description)

End Sub


Private Sub cmdFind_Click(Index As Integer)
    Dim oZipFind As PlusFind2.CZipFind

    On Error GoTo ErrHandler
    If Index = 0 Then '[1] 로고 찾기
        imgLogo.Picture = LoadPicture()
        
        dlgLogo.ShowOpen
        m_sPath = dlgLogo.FileName
        imgLogo.Picture = LoadPicture(m_sPath)
    ElseIf Index = 1 Then '[2] 주소찾기
    
    
        'S_201312_태을염직_99 에 의한 추가
        '위저드 우편번호  DB 정상 연결시
''        If g_bChkWizDBConn = False Then
''            g_bChkWizDBConn = PlusMDI.ConnectWizDB()
''        End If
        '위저드 우편번호  DB 정상 연결시
        If PlusMDI.ConnectWizDB() = False Then
            MsgBox "도로명 주소 DB연결 실패 !!!" & vbCrLf & "지속적인 연결 실패시 수동으로 입력하십시오.", vbCritical, "DB연결 실패"
            Exit Sub
        End If
        
    
        Set oZipFind = New PlusFind2.CZipFind
        'S_201312_태을염직_99 에 의한 수정(OLD: g_adoCon)
        oZipFind.Connection = g_adoWizCon
        
        
        'S_201312_태을염직_99 에 의한 추가
        If optOldNNew(0).Value = True Then      '도로명 주소
            oZipFind.Address1 = txtAddress1.Text
        Else                                    '지번 주소
        
            'S_201312_태을염직_99 에 의한 수정(OLD:oZipFind.Address1)
            oZipFind.AddressJiBun1 = txtName(7).Text
        End If
                    
    ''    oZipFind.Address1 = txtName(4)
        'S_201312_태을염직_99 에 의한 추가
        oZipFind.OldNNewSet = IIf(optOldNNew(0).Value = True, "0", "1")
    
    
''        'S_201110_대진텍스_03 에 의한 수정(OLD:4)
''        If Len(txtName(7)) > 0 Then oZipFind.Address1 = txtName(7)
        If oZipFind.Show() Then
            mskName(1) = oZipFind.ZipCode
''            txtName(7) = oZipFind.Address           '주소1

            'S_201312_태을염직_99 에 의한 수정-----------------------------------------------
            If oZipFind.OldNNewClss = "0" Then    '도로명 주소
                optOldNNew(0).Value = True
                
                txtAddress1.Text = oZipFind.Address
                txtAddress2.Text = oZipFind.AddressDetail
                txtAddressAssist.Text = oZipFind.AddressAssist
                txtGunMoolMngNo.Text = oZipFind.GunMoolMngNo
                
                txtAddress2.SetFocus
                
            Else
                optOldNNew(1).Value = True
                txtName(7).Text = oZipFind.Address
                'S_201110_대진텍스_03 에 의한 수정(OLD:5)
                txtName(8).SetFocus
            End If
            '----------------------------------------------------------------------------

        '2013.12.12 주석처리
''        Else
''            MsgBox LoadResString(252), vbInformation
        End If
    End If
    
    Set oZipFind = Nothing
    Exit Sub
    
ErrHandler:
    Set oZipFind = Nothing
    
    dlgLogo.FileName = ""
    imgLogo.Picture = LoadPicture()
End Sub

'S_201110_대진텍스_03 에 의한 추가
Private Sub cmdSave_Click()
    Dim rs As ADODB.Recordset
    Dim irow As Integer
   On Error GoTo ErrHandler
    
   ' If (QuestionBox("변경된 내용을 저장하시겠습니까?")) Then
   If (MsgBox("변경된 내용을 저장하시겠습니까?", vbYesNo + vbQuestion, "자료 저장") = vbYes) Then
        
       irow = grdData.Row           '현재 행 저장
       
        '데이터 체크
        If CheckData() = False Then Exit Sub
        
        Call SaveInfo 'db에 저장하는 부분...
        
'''        Call SaveRegistry  ' 레지스트리에 저장. 쓰지 않음.
        
        '-------------------------------------
        '업체정보 Get
        '-------------------------------------
        If Gf_DB_CM_GetCompanyInfo(rs, "Y") = True Then
    
            If rs.EOF = False Then
                g_companyInfo.Company_ID = Trim(CheckNull(rs!Company_ID))        '사업장ID
                g_companyInfo.Company_Name = Trim(CheckNull(rs!Company_Name))    '상호
                g_companyInfo.Chief = Trim(CheckNull(rs!Chief))                  '대표자명
                
                'S_201312_태을염직_99 에 의한 추가----------------------------------------
                g_companyInfo.OldNNewClss = Trim(CheckNull(rs!OldNNewClss))     '주소구분(0:도로명주소,1:지번주소)
                g_companyInfo.GunMoolMngNo = Trim(CheckNull(rs!GunMoolMngNo))   '건물고유식별코드
                g_companyInfo.Address1 = Trim(CheckNull(rs!Address1))           '도로명 기본주소
                g_companyInfo.Address2 = Trim(CheckNull(rs!Address2))           '도로명 상세주소
                g_companyInfo.AddressAssist = Trim(CheckNull(rs!AddressAssist)) '도로명 보조주소
                '----------------------------------------------------------------------------------
                
                g_companyInfo.AddressJiBun1 = Trim(CheckNull(rs!AddressJiBun1))            '지번주소1
                g_companyInfo.AddressJiBun2 = Trim(CheckNull(rs!AddressJiBun2))            '지번주소2
                g_companyInfo.Company_type = Trim(CheckNull(rs!Company_type))    '업태
                g_companyInfo.Category = Trim(CheckNull(rs!Category))            '업종
                g_companyInfo.Company_No = Trim(CheckNull(rs!Company_No))        '사업자번호
                
                '2012.02.27 추가- 거래명세서 출력을 위함
                g_companyInfo.Phone = Trim(CheckNull(rs!Phone))                 '전화번호
                g_companyInfo.Phone2 = Trim(CheckNull(rs!Phone2))               '전화번호2
                g_companyInfo.FaxNO = Trim(CheckNull(rs!FaxNO))                 '팩스번호
                
                '2013.02.04 추가
                g_companyInfo.BANK1 = Trim(CheckNull(rs!BANK1))                 '계좌번호1
                g_companyInfo.BANK2 = Trim(CheckNull(rs!BANK2))                 '계좌번호2
                g_companyInfo.BANK3 = Trim(CheckNull(rs!BANK3))                 '계좌번호3
                   
            End If
        End If
        
        Set rs = Nothing
        '-------------------------------------
        
''        grdData.Row = iRow              '데이터 재 조회
        Call FillGrid           '데이터 재 조회
        
    End If
    
    ''S_201110_대진텍스_03 에 의한 수정-주석처리
''    If Len(Trim(txtName(1))) > 0 Then
''        PlusMDI.Caption = LoadResString(101) & " - " & txtName(1)
''    Else
''        PlusMDI.Caption = LoadResString(101)
''    End If
''
''
''    Unload Me

    Exit Sub
    
    
ErrHandler:
    Set rs = Nothing
    Set oInfo = Nothing
    
    Call ErrorBox(Err.Number, "frmSetting.cmdSave_Click", Err.Description)
    
End Sub

Private Sub Form_Load()

    Me.Move 0, 0
    
    With dlgLogo
        .DialogTitle = "회사로고 열기"
        .Flags = cdlOFNHideReadOnly
        .Filter = "(*.bmp)|*.bmp"
        .CancelError = True
        
        .InitDir = App.Path
    End With
    
    Call SetOperate(Me)
    
    Call InitGrid
    Call FillGrid
    
    With cmdExit
        .MousePointer = ssCustom
        .MouseIcon = LoadResPicture("POINTER", vbResCursor)
        .Picture = LoadResPicture("EXIT", vbResIcon)
        .Cancel = True
    End With


    cmdSave.Picture = LoadResPicture("CHECK", vbResIcon)
    cmdCancel.Picture = LoadResPicture("CANCEL", vbResIcon)
        
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
''    imgTip.Picture = LoadResPicture("TIP", vbResIcon)
    cmdFind(0).Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind(1).Picture = LoadResPicture("FIND", vbResIcon)
End Sub

Private Sub InitGrid()
    Call SetVSFlexGrid(grdData)
    With grdData
        .Redraw = False
        
        
        .Rows = 1
        .Cols = 5
        
        .TextMatrix(0, 0) = "":             .ColWidth(0) = 450:     .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(0, 1) = "코드":         .ColWidth(1) = 0:       .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "상호":         .ColWidth(2) = 1200:     .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "사업자번호":   .ColWidth(3) = 1200:     .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(0, 4) = "대표자":       .ColWidth(4) = 450:     .ColAlignment(4) = flexAlignCenterCenter
        
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarHorizontal
        .Redraw = True
    End With
End Sub

Private Sub FillGrid()

    Dim oInfo As PlusLib2.CInfo
    Dim rs As ADODB.Recordset
    Dim sPath$, sFileName$
    Dim lsAdditemStr                    As String
    Dim lnsRows                         As Long
    
    On Error GoTo ErrLoadInfo:
 
    Set oInfo = New PlusLib2.CInfo
    oInfo.Connection = g_adoCon
    Set rs = oInfo.GetInfo(0, "", "")
    
    grdData.Rows = grdData.FixedRows
    
    Do While rs.EOF = False
        lnrows = lnrows + 1
        With grdData
            .Redraw = False
            
            lsAdditemStr = CStr(lnrows)                                                             '0)Row 수
            lsAdditemStr = lsAdditemStr & vbTab & Trim(CheckNull(rs!CompanyID))                      '1)자사코드
            lsAdditemStr = lsAdditemStr & vbTab & Trim(CheckNull(rs!KCompany))                       '2)상호
            lsAdditemStr = lsAdditemStr & vbTab & Format(CheckNull(rs!CompanyNo), "###-##-#####")    '3)사업자번호
            lsAdditemStr = lsAdditemStr & vbTab & Trim(CheckNull(rs!Chief))                          '4)대표자
            lsAdditemStr = lsAdditemStr & vbTab & Trim(CheckNull(rs!CompanyID))                      '5)자사코드

            .AddItem lsAdditemStr
                    
        
            .Redraw = True
        End With
        
        rs.MoveNext
    Loop
    
    If grdData.Rows > grdData.FixedRows Then
        grdData.Row = 1
    End If
    
    
    Exit Sub

ErrLoadInfo:
    Set rs = Nothing
    Set oInfo = Nothing
    
    Call ErrorBox(Err.Number, "frmSetting.FillGrid", Err.Description)
End Sub

Private Sub ShowData()
    Dim oInfo As PlusLib2.CInfo
    Dim rs As ADODB.Recordset
    Dim sPath$, sFileName$
    
    On Error GoTo ErrLoadInfo:
 
    Set oInfo = New PlusLib2.CInfo
    oInfo.Connection = g_adoCon
    
    '----------------------검색구분(0:조건없음,1:코드검색,2:상호검색),자사코드,자사상호
    Set rs = oInfo.GetInfo(1, grdData.TextMatrix(grdData.Row, 1), "")

    Do While rs.EOF = False
    
        '기존정보***********************************************************
        'S_201110_대진텍스_03 에 의한 수정-NEW
        txtName(0) = CheckNull(rs!CompanyID)        '[1] 자사코드
        txtName(1) = CheckNull(rs!KCompany)         '[2] 한글 상호
        txtName(2) = CheckNull(rs!ECompany)         '[3] 영문상호
        txtName(3) = CheckNull(rs!ShortCompany)     '[4] 약어
        txtName(4) = CheckNull(rs!Chief)            '[5] 대표자
        mskName(0) = CheckNull(rs!CompanyNo)        '[6] 사업자번호
        mskName(2) = CheckNull(rs!RegistID)         '[7] 주민/법인등록번호
        txtName(5) = CheckNull(rs!Condition)        '[8] 업태
        txtName(6) = CheckNull(rs!Category)         '[9] 업종
        mskName(1) = CheckNull(rs!ZipCode)          '[10] 우편번호
        
                
        'S_201312_태을염직_99 에 의한 추가-----------------------------------------------
        If CheckNull(rs!OldNNewClss) = "0" Then
            optOldNNew(0).Value = True     '도로명주소선택-수정할것
        Else
            optOldNNew(1).Value = True     '지번주소-수정할것
        End If
        txtGunMoolMngNo.Text = CheckNull(rs!GunMoolMngNo)       '건물관리 고유식별번호-수정할것

        txtAddress1.Text = CheckNull(rs!Address1)         '[11] 주소-도로명
        txtAddress2.Text = CheckNull(rs!Address2)          '[12] 주소2-도로명
        txtAddressAssist.Text = CheckNull(rs!AddressAssist)          '[12] 도로명 보조주소
        '-----------------------------------------------------------------------------------
        
        'S_201312_태을염직_99 에 의한 수정(OLD:rs!Address1)
        txtName(7) = CheckNull(rs!AddressJiBun1)         '[11] 주소-지번
        'S_201312_태을염직_99 에 의한 수정(OLD:rs!Address2)
        txtName(8) = CheckNull(rs!AddressJiBun2)         '[12] 주소2-지번
        txtName(9) = CheckNull(rs!Phone1)           '[13] 대표전화
        txtName(10) = CheckNull(rs!Phone2)          '[14] 전화번호
        txtName(11) = CheckNull(rs!FaxNO)           '[15] 팩스번호
        txtName(12) = CheckNull(rs!Email)           '[16] 이메일
        txtName(13) = CheckNull(rs!Homepage)        '[17] 홈페이지
        txtName(14) = CheckNull(rs!BANK1)           '[18] 계좌번호1
        txtName(15) = CheckNull(rs!BANK2)           '[19] 계좌번호2
        txtName(16) = CheckNull(rs!BANK3)           '[20] 계좌번호3
        
        '대표여부
        If CheckNull(rs!RPYn) = "Y" Then
            optRPYn(0).Value = True
            optRPYn(1).Value = False
            txtRPYN_OLD.Text = "Y"      '이전 대표여부 설정값
        Else
            optRPYn(0).Value = False
            optRPYn(1).Value = True
            txtRPYN_OLD.Text = "N"      '이전 대표여부 설정값
        End If

        '사용여부
        If CheckNull(rs!UseYn) = "Y" Then
            optUseYn(0).Value = True
            optUseYn(1).Value = False
        Else
            optUseYn(0).Value = False
            optUseYn(1).Value = True
        End If

        
        txtName(19) = g_sServer                    '[20] 서버 명->숨김
        txtName(20) = g_sDatabase                 '[21] DB 명->숨김
        '*****************************************************************
        
        '추가정보*********************************************************
        ' --WebPage로그인정보
        txtWebData(0).Text = CheckNull(rs!WebPortFrom)              'WebPage포트From
        txtWebData(1).Text = CheckNull(rs!WebPortTo)                'WebPage포트To
        txtWebData(2).Text = CheckNull(rs!WebID1)                   'WebPage로그인ID1
        txtWebData(3).Text = CheckNull(rs!WebPass1)                 'WebPage로그인암호1
        txtWebData(4).Text = CheckNull(rs!WebAuthCode1)             'WebPage로그인인증코드1
        txtWebData(5).Text = CheckNull(rs!WebID2)                   'WebPage로그인ID2
        txtWebData(6).Text = CheckNull(rs!WebPass2)                 'WebPage로그인암호2
        txtWebData(7).Text = CheckNull(rs!WebAuthCode2)             'WebPage로그인인증코드2
        '
        ' --FTP로그인정보
        txtFTPData(0).Text = CheckNull(rs!FTPPage)                  'FTP주소
        txtFTPData(1).Text = CheckNull(rs!FTPPortFrom)              'FTP포트From
        txtFTPData(2).Text = CheckNull(rs!FTPPortTo)                'FTP포트To
        txtFTPData(3).Text = CheckNull(rs!FTPID1)                   'FTP로그인ID1
        txtFTPData(4).Text = CheckNull(rs!FTPPass1)                 'FTP로그인암호1
        txtFTPData(5).Text = CheckNull(rs!FTPAuthCode1)             'FTP로그인인증코드1
        txtFTPData(6).Text = CheckNull(rs!FTPID2)                   'FTP로그인ID2
        txtFTPData(7).Text = CheckNull(rs!FTPPass2)                 'FTP로그인암호2
        txtFTPData(8).Text = CheckNull(rs!FTPAuthCode2)             'FTP로그인인증코드2
        '
        ' --SMS서버1그인정보
        txtSMS1Data(0).Text = CheckNull(rs!SMSURL1)                 '문자전송서버1주소
        txtSMS1Data(1).Text = CheckNull(rs!SMSPortFrom1)            '문자전송서버1포트From
        txtSMS1Data(2).Text = CheckNull(rs!SMSPortTo1)              '문자전송서버1포트To
        txtSMS1Data(3).Text = CheckNull(rs!SMSID1)                  '문자전송서버1아이디
        txtSMS1Data(4).Text = CheckNull(rs!SMSPASS1)                '문자전송서버1암호
        txtSMS1Data(5).Text = CheckNull(rs!SMSAuthCode1)            '문자전송서버1인증코드
        '
        ' --SMS서버2로그인정보
        txtSMS2Data(0).Text = CheckNull(rs!SMSURL2)                 '문자전송서버2주소
        txtSMS2Data(1).Text = CheckNull(rs!SMSPortFrom2)            '문자전송서버2포트From
        txtSMS2Data(2).Text = CheckNull(rs!SMSPortTo2)              '문자전송서버2포트To
        txtSMS2Data(3).Text = CheckNull(rs!SMSID2)                  '문자전송서버2아이디
        txtSMS2Data(4).Text = CheckNull(rs!SMSPASS2)                '문자전송서버2암호
        txtSMS2Data(5).Text = CheckNull(rs!SMSAuthCode2)            '문자전송서버2인증코드
        '*****************************************************************
        
        rs.MoveNext
    Loop

    sPath = App.Path & "\"
    sFileName = "Logo.bmp"
    If oInfo.GetCompanyLogo(sPath, sFileName) Then
        imgLogo.Picture = LoadPicture(sPath & sFileName)
    Else
        imgLogo.Picture = LoadPicture()
    End If
    
    rs.Close
    Set rs = Nothing
    Set oInfo = Nothing
    Exit Sub
    
ErrLoadInfo:
    Set rs = Nothing
    Set oInfo = Nothing
    Call ErrorBox(Err.Number, "frmSetting.LoadInfo", Err.Description)
End Sub


'S_201110_대진텍스_03 에 의한 추 가
Private Sub grdData_DblClick()
    '에디트 true
    
    If grdData.Rows = grdData.FixedRows Or grdData.Row <= 0 Then Exit Sub
    
    m_sFlag = ID_UPDATE
    Call ChangeMode(Me, False)
    
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    
    pnlMsg.Caption = LoadResString(303)
    txtName(0).Locked = True
    txtName(1).SetFocus
    
            
    '대표여부로 설정 되었으면
    If txtRPYN_OLD.Text = "Y" Then
        fraRPYN.Enabled = False
        fraUseYN.Enabled = False
    Else
        fraRPYN.Enabled = True
        fraUseYN.Enabled = True
    End If
    
    'S_201312_태을염직_99 에 의한 추가-----------------------------------------------
    If optOldNNew(0).Value = True Then
        fraDoro.Enabled = True
        fraJiBun.Enabled = False
    Else
        fraDoro.Enabled = False
        fraJiBun.Enabled = True
    End If
    '-------------------------------------------------------------------------
            
End Sub

'S_201110_대진텍스_03 에 의한 추가
Private Sub grdData_RowColChange()
    Call ShowData
End Sub

'S_201110_대진텍스_03 에 의한 추가
Private Function CheckData() As Boolean
    Dim i%
    CheckData = True
    If m_sFlag = ID_ADDNEW Or m_sFlag = ID_UPDATE Then
        '상호체크
        If Len(Trim(txtName(1))) = 0 Then
            MsgBox "상호가 입력되지 않았습니다.", vbInformation
            txtName(1).SetFocus
            CheckData = False
            Exit Function
        End If
        
        '대표자 체크
        If Len(Trim(txtName(4))) = 0 Then
            MsgBox "대표자가 입력되지 않았습니다.", vbInformation
            txtName(4).SetFocus
            CheckData = False
            Exit Function
        End If
        
        '사업자등록번호 체크
        If Len(Trim(mskName(0))) < 10 Then
            MsgBox "사업자등록번호가 정상적이지 않거나 입력되지 않았습니다.", vbInformation
            mskName(0).SetFocus
            CheckData = False
            Exit Function
        End If
        
        '법인/주민등록번호 체크-입력시에만 체크
        If Len(Trim(mskName(2))) > 0 And Len(Trim(mskName(2))) < 13 Then
            MsgBox "주민/법인등록번호가 정상적이지 않았습니다.", vbInformation
            mskName(2).SetFocus
            CheckData = False
            Exit Function
        End If
        
        
        '주소 체크
        'S_201312_태을염직_99 에 의한 수정-txtAddress1 조건 추가
        If (optOldNNew(0).Value = True And Len(Trim(txtAddress1.Text)) = 0) _
            Or (optOldNNew(1).Value = True And Len(Trim(txtName(7))) = 0) Then
            MsgBox "주소가 입력되지 않았습니다.", vbInformation
            
            If optOldNNew(0).Value = True Then      '도로명 주소
                txtAddress1.SetFocus
            Else
                txtName(7).SetFocus
            End If
            
            CheckData = False
            Exit Function
        End If
        
                
        '대표사용이면서 사용 안함인지 체크
        If optRPYn(0).Value = True And optUseYn(1).Value = True Then
            MsgBox "대표사용여부로 체크시 사용여부를 [예]로 설정하십시오.", vbInformation
            CheckData = False
            Exit Function
        End If
        
    End If
End Function

Private Sub mskName_GotFocus(Index As Integer)
    With mskName(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub mskName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub mskName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call NextFocus
    End If
End Sub

'S_201312_태을염직_99 에 의한 추가
Private Sub optOldNNew_Click(Index As Integer)
    If optOldNNew(0).Value = True Then
        fraDoro.Enabled = True
        fraJiBun.Enabled = False
    Else
        fraDoro.Enabled = False
        fraJiBun.Enabled = True
    End If
End Sub

Private Sub optRPYn_Click(Index As Integer)
    '대표사용부 체크시 사용여부 프레임은 사용 못함
    If optRPYn(0).Value = True Then
        optUseYn(0).Value = True
        optUseYn(1).Value = False
        fraUseYN.Enabled = False
    Else
        fraUseYN.Enabled = True

    End If
End Sub

'S_201312_태을염직_99 에 의한 추가
Private Sub txtAddress1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then

        Call cmdFind_Click(1)

    End If
End Sub

Private Sub txtName_GotFocus(Index As Integer)
    With txtName(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 7 Then
            Call cmdFind_Click(1)
        End If
        
        Call NextFocus
    End If
End Sub
