VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmDyePattern 
   BackColor       =   &H00000000&
   BorderStyle     =   0  '없음
   ClientHeight    =   8955
   ClientLeft      =   210
   ClientTop       =   1590
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMode 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Left            =   7620
      TabIndex        =   40
      Top             =   150
      Width           =   5115
      Begin VB.CommandButton cmdExit 
         Height          =   675
         Left            =   4440
         MousePointer    =   99  '사용자 정의
         Style           =   1  '그래픽
         TabIndex        =   52
         Top             =   150
         Width           =   645
      End
      Begin VSFlex7LCtl.VSFlexGrid vsMode 
         Height          =   2865
         Left            =   30
         TabIndex        =   51
         Top             =   840
         Width           =   5040
         _cx             =   8890
         _cy             =   5054
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483642
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  '투명
         Caption         =   "모드 리스트"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   90
         TabIndex        =   48
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "온도( 0 - 9)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   0
         Left            =   1230
         TabIndex        =   47
         Top             =   210
         Width           =   960
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "급수(10-19)"
         ForeColor       =   &H000080FF&
         Height          =   180
         Index           =   1
         Left            =   2310
         TabIndex        =   46
         Top             =   210
         Width           =   960
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "배수(20-29)"
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Index           =   2
         Left            =   3420
         TabIndex        =   45
         Top             =   210
         Width           =   960
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "준비(30-39)"
         ForeColor       =   &H0000C000&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   44
         Top             =   570
         Width           =   960
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "보조(40-49)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   4
         Left            =   1230
         TabIndex        =   43
         Top             =   570
         Width           =   960
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "투입(50-59)"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   5
         Left            =   2310
         TabIndex        =   42
         Top             =   570
         Width           =   960
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "기타(60-69)"
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Index           =   6
         Left            =   3420
         TabIndex        =   41
         Top             =   570
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C00000&
         BackStyle       =   1  '투명하지 않음
         Height          =   240
         Index           =   2
         Left            =   30
         Shape           =   4  '둥근 사각형
         Top             =   150
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Left            =   2820
      TabIndex        =   26
      Top             =   150
      Width           =   4815
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   34
         Top             =   3420
         Width           =   765
      End
      Begin VB.TextBox txtPTName 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   33
         Top             =   150
         Width           =   2355
      End
      Begin VB.TextBox txtPTNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1410
         MaxLength       =   3
         TabIndex        =   32
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "행삽입"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   31
         Top             =   3420
         Width           =   765
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "행삭제"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   795
         TabIndex        =   30
         Top             =   3420
         Width           =   765
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2550
         TabIndex        =   29
         Top             =   3420
         Width           =   735
      End
      Begin VB.CommandButton cmdKill 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3285
         TabIndex        =   28
         Top             =   3420
         Width           =   735
      End
      Begin VB.CommandButton cmdNewPt 
         Caption         =   "신규"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4020
         TabIndex        =   27
         Top             =   3420
         Width           =   735
      End
      Begin VSFlex7LCtl.VSFlexGrid vsPatternEdit 
         Height          =   2715
         Left            =   30
         TabIndex        =   50
         Top             =   420
         Width           =   4740
         _cx             =   8361
         _cy             =   4789
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
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   5
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         Editable        =   2
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  '투명
         Caption         =   "공정 편집"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   60
         TabIndex        =   39
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         TabIndex        =   38
         Top             =   210
         Width           =   390
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H00000000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "메시지"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   37
         Top             =   3150
         Width           =   3795
      End
      Begin VB.Label Label9 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "Message"
         Height          =   180
         Left            =   105
         TabIndex        =   36
         Top             =   3180
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "명칭"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   2010
         TabIndex        =   35
         Top             =   210
         Width           =   390
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C00000&
         BackStyle       =   1  '투명하지 않음
         Height          =   240
         Index           =   0
         Left            =   30
         Shape           =   4  '둥근 사각형
         Top             =   150
         Width           =   930
      End
   End
   Begin VB.PictureBox picDyeConGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   4875
      Left            =   150
      MouseIcon       =   "frmDyePattern.frx":0000
      MousePointer    =   2  '십자형
      ScaleHeight     =   321
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   834
      TabIndex        =   0
      Top             =   3900
      Width           =   12570
      Begin VB.Label lblIntervalTemper 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "30"
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   23
         Top             =   3630
         Width           =   180
      End
      Begin VB.Label lblIntervalTemper 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "60"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   22
         Top             =   2730
         Width           =   180
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   5
         X1              =   50
         X2              =   950
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   4
         X1              =   50
         X2              =   950
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   50
         X2              =   50
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   50
         X2              =   950
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   1
         X1              =   50
         X2              =   950
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   2
         X1              =   50
         X2              =   950
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   3
         X1              =   50
         X2              =   652
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   1
         X1              =   110
         X2              =   110
         Y1              =   2
         Y2              =   302
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   2
         X1              =   170
         X2              =   170
         Y1              =   2
         Y2              =   302
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   3
         X1              =   230
         X2              =   230
         Y1              =   2
         Y2              =   302
      End
      Begin VB.Label lblBase 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0"
         Height          =   180
         Left            =   630
         TabIndex        =   21
         Top             =   4530
         Width           =   90
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "20"
         Height          =   180
         Index           =   0
         Left            =   1560
         TabIndex        =   20
         Top             =   4530
         Width           =   180
      End
      Begin VB.Label lblIntervalTemper 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "90"
         Height          =   180
         Index           =   2
         Left            =   510
         TabIndex        =   19
         Top             =   1830
         Width           =   180
      End
      Begin VB.Label lblIntervalTemper 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "120"
         Height          =   180
         Index           =   3
         Left            =   450
         TabIndex        =   18
         Top             =   930
         Width           =   270
      End
      Begin VB.Label lblIntervalTemper 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "150"
         Height          =   180
         Index           =   4
         Left            =   450
         TabIndex        =   17
         Top             =   30
         Width           =   285
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "온도"
         Height          =   180
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   360
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "분"
         Height          =   180
         Left            =   12330
         TabIndex        =   15
         Top             =   4530
         Width           =   180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   8
         X1              =   290
         X2              =   290
         Y1              =   2
         Y2              =   302
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   9
         X1              =   350
         X2              =   350
         Y1              =   0
         Y2              =   302
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   10
         X1              =   410
         X2              =   410
         Y1              =   2
         Y2              =   302
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   11
         X1              =   530
         X2              =   530
         Y1              =   2
         Y2              =   302
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   12
         X1              =   470
         X2              =   470
         Y1              =   2
         Y2              =   302
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   13
         X1              =   650
         X2              =   650
         Y1              =   2
         Y2              =   302
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   14
         X1              =   590
         X2              =   590
         Y1              =   2
         Y2              =   302
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "40"
         Height          =   180
         Index           =   1
         Left            =   2460
         TabIndex        =   14
         Top             =   4530
         Width           =   180
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "60"
         Height          =   180
         Index           =   2
         Left            =   3360
         TabIndex        =   13
         Top             =   4530
         Width           =   180
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "80"
         Height          =   180
         Index           =   3
         Left            =   4260
         TabIndex        =   12
         Top             =   4530
         Width           =   180
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "100"
         Height          =   180
         Index           =   4
         Left            =   5130
         TabIndex        =   11
         Top             =   4530
         Width           =   270
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "120"
         Height          =   180
         Index           =   5
         Left            =   6030
         TabIndex        =   10
         Top             =   4530
         Width           =   270
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "140"
         Height          =   180
         Index           =   6
         Left            =   6930
         TabIndex        =   9
         Top             =   4530
         Width           =   270
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "160"
         Height          =   180
         Index           =   7
         Left            =   7830
         TabIndex        =   8
         Top             =   4530
         Width           =   270
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "180"
         Height          =   180
         Index           =   8
         Left            =   8730
         TabIndex        =   7
         Top             =   4530
         Width           =   270
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "200"
         Height          =   180
         Index           =   9
         Left            =   9600
         TabIndex        =   6
         Top             =   4530
         Width           =   270
      End
      Begin VB.Line linCurrTime 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  '점
         BorderWidth     =   2
         DrawMode        =   9  '마스크 펜이 아님
         Visible         =   0   'False
         X1              =   24
         X2              =   24
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   4
         X1              =   710
         X2              =   710
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   5
         X1              =   770
         X2              =   770
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   6
         X1              =   830
         X2              =   830
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   7
         X1              =   890
         X2              =   890
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '점
         Index           =   15
         X1              =   950
         X2              =   950
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "220"
         Height          =   180
         Index           =   10
         Left            =   10500
         TabIndex        =   5
         Top             =   4530
         Width           =   270
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "240"
         Height          =   180
         Index           =   11
         Left            =   11400
         TabIndex        =   4
         Top             =   4530
         Width           =   270
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "260"
         Height          =   180
         Index           =   12
         Left            =   12060
         TabIndex        =   3
         Top             =   4530
         Width           =   270
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "280"
         Height          =   180
         Index           =   13
         Left            =   13200
         TabIndex        =   2
         Top             =   4530
         Width           =   270
      End
      Begin VB.Label lblIntervalTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "300"
         Height          =   180
         Index           =   14
         Left            =   13800
         TabIndex        =   1
         Top             =   4530
         Width           =   270
      End
   End
   Begin VB.Frame fraPatternList 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Left            =   150
      TabIndex        =   24
      Top             =   150
      Width           =   2655
      Begin VSFlex7LCtl.VSFlexGrid vsPatternList 
         Height          =   3285
         Left            =   30
         TabIndex        =   49
         Top             =   420
         Width           =   2580
         _cx             =   4551
         _cy             =   5794
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  '투명
         Caption         =   "염색패턴 리스트"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   570
         TabIndex        =   25
         Top             =   180
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C00000&
         BackStyle       =   1  '투명하지 않음
         Height          =   240
         Index           =   1
         Left            =   60
         Shape           =   4  '둥근 사각형
         Top             =   150
         Width           =   2520
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   4
      FillColor       =   &H00FFFFFF&
      Height          =   8895
      Left            =   30
      Top             =   30
      Width           =   12825
   End
End
Attribute VB_Name = "frmDyePattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gXInterval As Integer
Private gYInterval As Integer

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub Form_Activate()
    cmdExit.Picture = LoadResPicture("EXIT", vbResIcon)
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    lblMsg = ""
    Call vsPatternSet
    Call vsPatternEditSet
    Call vsModeSet
    Call vsModeFill
    Call vsPatternFill
    txtPTNo = ""
    txtPTName = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub SettingInterval()
    gXInterval = CInt((Line2(0).x2 - Line2(0).x1) / CInt(lblIntervalTime.Count) / CInt(lblIntervalTime(0).Caption))
    gYInterval = CInt((Line1(0).y2 - Line1(0).y1) / CInt(lblIntervalTemper.Count) / CInt(lblIntervalTemper(0).Caption))
End Sub

Private Sub vsModeFill()
    Dim oDyeCon As PlusLib2.CDyeCon
    Dim tRs As Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    Set oDyeCon = New PlusLib2.CDyeCon
    oDyeCon.Connection = g_adoCon
    oDyeCon.UserName = g_sUserName

    Set tRs = oDyeCon.GetMode()
    With vsMode
        If tRs.RecordCount > 0 Then
            For i = 1 To tRs.RecordCount
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = tRs!ModeNo
                .TextMatrix(i, 1) = tRs!ModeName
                If Trim(tRs!selectmsg1) = "" Then
                    .TextMatrix(i, 2) = ""
                    .TextMatrix(i, 3) = ""
                    .TextMatrix(i, 4) = ""
                Else
                    .TextMatrix(i, 2) = tRs!selectmsg1
                    .TextMatrix(i, 3) = tRs!selectmin1
                    .TextMatrix(i, 4) = tRs!selectmax1
                End If
                .TextMatrix(i, 5) = tRs!selectunit1
                If Trim(tRs!selectmsg2) = "" Then
                    .TextMatrix(i, 6) = ""
                    .TextMatrix(i, 7) = ""
                    .TextMatrix(i, 8) = ""
                Else
                    .TextMatrix(i, 6) = tRs!selectmsg2
                    .TextMatrix(i, 7) = tRs!selectmin2
                    .TextMatrix(i, 8) = tRs!selectmax2
                End If
                .TextMatrix(i, 9) = tRs!selectunit2
                
                If i < 11 Then
                    .Cell(flexcpBackColor, i, 0, i, 0) = &HC0C0FF
                ElseIf i < 21 Then
                    .Cell(flexcpBackColor, i, 0, i, 0) = &HC0E0FF
                ElseIf i < 31 Then
                    .Cell(flexcpBackColor, i, 0, i, 0) = &HC0FFFF
                ElseIf i < 41 Then
                    .Cell(flexcpBackColor, i, 0, i, 0) = &HC0FFC0
                ElseIf i < 51 Then
                    .Cell(flexcpBackColor, i, 0, i, 0) = &HFFC0C0
                ElseIf i < 61 Then
                    .Cell(flexcpBackColor, i, 0, i, 0) = &HFF8080
                Else
                    .Cell(flexcpBackColor, i, 0, i, 0) = &HFFC0FF
                End If
                
                tRs.MoveNext
            Next i
        End If
    End With
    tRs.Close
    Set tRs = Nothing
    Set oDyeCon = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    Set tRs = Nothing
    Set oDyeCon = Nothing
    Call ErrorBox(Err.Number, "frmDyePattern.vsModeFill", Err.Description)
End Sub

Private Sub vsPatternFill()
    Dim oDyeCon As PlusLib2.CDyeCon
    Dim sRs As Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    Set oDyeCon = New PlusLib2.CDyeCon
    oDyeCon.Connection = g_adoCon
    oDyeCon.UserName = g_sUserName

    Set sRs = oDyeCon.GetPatternGroup(1, 0, 0)  ' 1: Rapid, 0: 0번호기, 0: 0번 패턴
    
    With vsPatternList
        If sRs.RecordCount > 0 Then
            For i = 1 To sRs.RecordCount
                .Rows = .Rows + 1
            
                .TextMatrix(i, 0) = sRs!PtNo
                .TextMatrix(i, 1) = sRs!PtName
                sRs.MoveNext
            Next i
        End If
    End With
    sRs.Close
    Set sRs = Nothing
    Set oDyeCon = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    Set sRs = Nothing
    Set oDyeCon = Nothing
    Call ErrorBox(Err.Number, "frmDyePattern.vsPatternFill", Err.Description)
End Sub

Private Sub vsPatternList_Click()
    Dim oDyeCon As PlusLib2.CDyeCon
    Dim tRs As ADODB.Recordset
    Dim i%

    On Error GoTo ErrHandler

    If vsPatternList.Rows > 1 And vsPatternList.Row >= vsPatternList.FixedRows Then
        Set oDyeCon = New PlusLib2.CDyeCon
        oDyeCon.Connection = g_adoCon
        oDyeCon.UserName = g_sUserName
        
        txtPTNo = vsPatternList.TextMatrix(vsPatternList.Row, 0)
        txtPTName = vsPatternList.TextMatrix(vsPatternList.Row, 1)
        
        Set tRs = oDyeCon.GetPattern(1, 0, CInt(vsPatternList.TextMatrix(vsPatternList.Row, 0)))
        If tRs.RecordCount > 0 Then
            With vsPatternEdit
                .Rows = tRs.RecordCount + 1
                For i = 1 To tRs.RecordCount
                    .TextMatrix(i, 1) = tRs!worksection
                    .TextMatrix(i, 2) = tRs!ModeNo
                    .TextMatrix(i, 3) = tRs!ModeName
                    If tRs!SelNo1 = 0 Then
                        .TextMatrix(i, 4) = ""
                        .TextMatrix(i, 5) = ""
                        .TextMatrix(i, 6) = ""
                    Else
                        .TextMatrix(i, 4) = tRs!SelNo1
                        Select Case CStr(tRs!ModeNo)
                            Case "13", "14", "15":
                                Select Case CStr(tRs!SelNo1)
                                    Case "1":   .TextMatrix(i, 5) = "초기"
                                    Case "2":   .TextMatrix(i, 5) = "수세1"
                                    Case "3":   .TextMatrix(i, 5) = "수세2"
                                    Case "4":   .TextMatrix(i, 5) = "염색"
                                End Select
                            Case "17":
                                Select Case CStr(tRs!SelNo1)
                                    Case "1":   .TextMatrix(i, 9) = "L"
                                                .TextMatrix(i, 5) = "수량"
                                    Case "2":   .TextMatrix(i, 9) = "분"
                                                .TextMatrix(i, 5) = "시간"
                                    Case "3":   .TextMatrix(i, 9) = "도"
                                                .TextMatrix(i, 5) = "온도"
                                    Case Else:
                                End Select
                            Case "18":
                                Select Case CStr(tRs!SelNo1)
                                    Case "1":   .TextMatrix(i, 9) = "L"
                                                .TextMatrix(i, 5) = "수량"
                                    Case "2":   .TextMatrix(i, 9) = "분"
                                                .TextMatrix(i, 5) = "시간"
                                    Case "3":   .TextMatrix(i, 9) = "도"
                                                .TextMatrix(i, 5) = "온도"
                                    Case Else:
                                End Select
                            Case "19":
                                Select Case CStr(tRs!SelNo1)
                                    Case "1":   .TextMatrix(i, 9) = "회"
                                                .TextMatrix(i, 5) = "회수"
                                    Case "2":   .TextMatrix(i, 9) = "분"
                                                .TextMatrix(i, 5) = "시간"
                                    Case "3":   .TextMatrix(i, 9) = "도"
                                                .TextMatrix(i, 5) = "온도"
                                    Case Else:
                                End Select
                            Case "40":
                                Select Case CStr(tRs!SelNo1)
                                    Case "1":   .TextMatrix(i, 9) = "도"
                                                .TextMatrix(i, 5) = "초기"
                                    Case "2":   .TextMatrix(i, 9) = ""
                                                .TextMatrix(i, 5) = "수세1"
                                    Case "3":   .TextMatrix(i, 9) = ""
                                                .TextMatrix(i, 5) = "수세2"
                                    Case "4":   .TextMatrix(i, 9) = ""
                                                .TextMatrix(i, 5) = "염색"
                                    Case Else:
                                End Select
                            Case "41", "42":
                                Select Case CStr(tRs!SelNo1)
                                    Case "1", "2", "3": .TextMatrix(i, 9) = "도"
                                                        .TextMatrix(i, 5) = "급수"
                                    Case "4", "5", "6": .TextMatrix(i, 9) = ""
                                                        .TextMatrix(i, 5) = "리턴"
                                    Case Else:
                                End Select
                            Case Else:  .TextMatrix(i, 5) = tRs!SelNo1
                                        .TextMatrix(i, 9) = tRs!selectunit2
                        End Select
                        .TextMatrix(i, 6) = tRs!selectunit1
                    End If
                    If tRs!SelNo2 = 0 Then
                        .TextMatrix(i, 7) = ""
                        .TextMatrix(i, 8) = ""
                        .TextMatrix(i, 9) = ""
                    Else
                        .TextMatrix(i, 7) = tRs!SelNo2
                        .TextMatrix(i, 8) = tRs!SelNo2
                        Select Case tRs!ModeNo
                            Case 13, 14, 15, 17, 18, 19, 40, 41, 42
                            Case Else:
                                .TextMatrix(i, 9) = tRs!selectunit2
                        End Select
                    End If
                    
                    tRs.MoveNext
                Next i
            End With
            Call SettingInterval
            Call GraphRefresh
        End If
        tRs.Close
        Set tRs = Nothing
        Set oDyeCon = Nothing
        Screen.MousePointer = vbDefault
    End If
    
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    Set tRs = Nothing
    Set oDyeCon = Nothing
    Call ErrorBox(Err.Number, "frmDyePattern.vsPatternList_Click", Err.Description)
End Sub

Private Sub cmdNewPt_Click()
    txtPTNo.Enabled = True
End Sub

Private Sub GraphRefresh()
Dim x1 As Integer
Dim y1 As Integer
Dim x2 As Integer
Dim y2 As Integer
Dim iRow As Integer

    With vsPatternEdit
        If .Rows > .FixedRows + 1 Then
            x1 = 50
            y1 = 300
            picDyeConGraph.Cls
            picDyeConGraph.DrawWidth = 3
            
            For iRow = 1 To .Rows - 1
                If Trim(.TextMatrix(iRow, 2)) <> "" Then
                    If CInt(.TextMatrix(iRow, 2)) < 6 Then
                    
                        Select Case CInt(.TextMatrix(iRow, 2))
                            Case 0: '시작온도
                                x2 = x1
                                y2 = 300 - (CInt(.TextMatrix(iRow, 4)) * gYInterval)
                                picDyeConGraph.Line (x1, y1)-(x2, y2), RGB(0, 0, 255)
                            Case 1: '진행
                                x2 = x1
                                y2 = 300
                                picDyeConGraph.Line (x1, y1)-(x1, y2), RGB(0, 0, 255)
                                x1 = x2
                                y1 = y2
                                x2 = x1 + (CInt(.TextMatrix(iRow, 7)) * gXInterval)
                                y2 = 300
                                picDyeConGraph.Line (x1, y1)-(x2, y2), RGB(0, 0, 255)
                            Case 2: '승온
                                x2 = x1 + (CInt("0" & .TextMatrix(iRow, 7)) * gXInterval)
                                y2 = 300 - (CInt(.TextMatrix(iRow, 4)) * gYInterval)
                                picDyeConGraph.Line (x1, y1)-(x2, y2), RGB(0, 0, 255)
                            Case 3: '냉각
                                x2 = x1 + (CInt(.TextMatrix(iRow, 7)) * gXInterval)
                                y2 = 300 - (CInt("0" & .TextMatrix(iRow, 4)) * gYInterval)
                                picDyeConGraph.Line (x1, y1)-(x2, y2), RGB(0, 0, 255)
                            Case 4: '유지1(승온유지)
                                x2 = x1 + (CInt(.TextMatrix(iRow, 7)) * gXInterval)
                                y2 = y1
                                picDyeConGraph.Line (x1, y1)-(x2, y2), RGB(0, 0, 255)
                            Case 5: '유지2(승온/냉각)
                                x2 = x1 + (CInt(.TextMatrix(iRow, 7)) * gXInterval)
                                y2 = y1
                                picDyeConGraph.Line (x1, y1)-(x2, y2), RGB(0, 0, 255)
                        End Select
                        x1 = x2
                        y1 = y2
                    End If
                End If
            Next iRow
        End If
    End With
End Sub

Private Sub cmdAdd_Click()
Dim i As Integer
Dim j As Integer
Dim iTempRow As Integer

    With vsPatternEdit
        If .Row >= .FixedRows Then
            iTempRow = .Row
            .Rows = .Rows + 1
            For i = .Rows - 1 To iTempRow + 1 Step -1
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = .TextMatrix(i - 1, j)
                Next j
            Next i
            .Cell(flexcpText, iTempRow, 0, iTempRow, .Cols - 1) = ""
        End If
    End With
End Sub

Private Sub cmdClear_Click()
    Call vsPatternEditSet
'    txtDyeID = ""
    txtPTNo = ""
    txtPTName = ""
    picDyeConGraph.Cls
End Sub

Private Sub cmdDel_Click()
Dim i As Integer
Dim j As Integer
Dim iTempRow As Integer

    With vsPatternEdit
        If .Row >= .FixedRows And .Rows > 2 Then
            iTempRow = .Row
            For i = iTempRow + 1 To .Rows - 1
                For j = 0 To .Cols - 1
                    .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                Next j
            Next i

            .Rows = .Rows - 1

            Call Seq_Sort
        End If
    End With
    Call SettingInterval
    Call GraphRefresh

End Sub

Private Sub Seq_Sort()
Dim iRow As Integer
Dim sTemp As String

    With vsPatternEdit
        sTemp = Trim(.TextMatrix(1, 1))
        For iRow = 1 To .Rows - 1
            If Trim(.TextMatrix(iRow, 1)) <> "" Then
                If .TextMatrix(iRow, 1) = sTemp Then
                    .TextMatrix(iRow, 1) = sTemp
                Else
                    .TextMatrix(iRow, 1) = CStr(CInt("0" & sTemp) + 1)
                    sTemp = CStr(CInt("0" & sTemp) + 1)
                End If
            End If
        Next iRow
    End With
End Sub

Private Sub vsPatternSet()
    Dim iCol As Integer

    With vsPatternList
        .Rows = 1:      .Cols = 2
        .TextMatrix(0, 0) = "번호":         .ColWidth(0) = 500:     .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(0, 1) = "공정 명":      .ColWidth(1) = 2000
        
        For iCol = 0 To .Cols - 1
            .FixedAlignment(iCol) = flexAlignCenterCenter
        Next iCol
    End With
End Sub

Private Sub lblMode_DblClick(Index As Integer)
    With vsMode
        Select Case Index
            Case 0:     .TopRow = 1
            Case 1:     .TopRow = 11
            Case 2:     .TopRow = 21
            Case 3:     .TopRow = 31
            Case 4:     .TopRow = 41
            Case 5:     .TopRow = 51
            Case 6:     .TopRow = 61
        End Select
    End With
End Sub

Private Sub vsPatternEditSet()
Dim iCol As Integer

    With vsPatternEdit
        .Rows = 2:      .Cols = 10
        
        .TextMatrix(0, 0) = " ":            .ColWidth(0) = 0
        .TextMatrix(0, 1) = "구간":         .ColWidth(1) = 500:     .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "모드번호":     .ColWidth(2) = 0
        .TextMatrix(0, 3) = "작업명":       .ColWidth(3) = 1600
        .TextMatrix(0, 4) = "선택1값":      .ColWidth(4) = 0
        .TextMatrix(0, 5) = "선택1":        .ColWidth(5) = 600
        .TextMatrix(0, 6) = "선택1":        .ColWidth(6) = 600
        .TextMatrix(0, 7) = "선택2값":      .ColWidth(7) = 0
        .TextMatrix(0, 8) = "선택2":        .ColWidth(8) = 600
        .TextMatrix(0, 9) = "선택2":        .ColWidth(9) = 600
        .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        
        For iCol = 0 To .Cols - 1
            .FixedAlignment(iCol) = flexAlignCenterCenter
        Next iCol
    End With
End Sub

Private Sub vsModeSet()
Dim iCol As Integer

    With vsMode
        .Rows = 1:          .Cols = 11
        
        .TextMatrix(0, 0) = "":             .ColWidth(0) = 350
        .TextMatrix(0, 1) = "모드 명칭":    .ColWidth(1) = 1600
        .TextMatrix(0, 2) = "선택1":    .ColWidth(2) = 800
        .TextMatrix(0, 3) = "선택1min":     .ColWidth(3) = 0
        .TextMatrix(0, 4) = "선택1max":     .ColWidth(4) = 0
        .TextMatrix(0, 5) = "단위":         .ColWidth(5) = 500
        .TextMatrix(0, 6) = "선택2":    .ColWidth(6) = 800
        .TextMatrix(0, 7) = "선택2min":     .ColWidth(7) = 0
        .TextMatrix(0, 8) = "선택2max":     .ColWidth(8) = 0
        .TextMatrix(0, 9) = "단위":         .ColWidth(9) = 800
        .TextMatrix(0, 10) = "모드그룹":     .ColWidth(10) = 0
        
        For iCol = 0 To .Cols - 1
            .FixedAlignment(iCol) = flexAlignCenterCenter
        Next iCol
    End With
End Sub



Private Sub picDyeConGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iTemp As Integer
    Dim iTime As Integer
On Error Resume Next
    
    If X >= Line2(0).x1 And X <= Line2(0).x2 And Y >= Line1(0).y1 And Y <= Line1(0).y2 Then
        iTemp = (300 - Y) \ gYInterval
        iTime = (X - 50) \ gXInterval
        picDyeConGraph.ToolTipText = " 온도:" & iTemp & "(℃), " & "시간:" & iTime & "(분)"
    Else
        picDyeConGraph.ToolTipText = ""
    End If
End Sub

Private Sub txtDyeID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPTName_LostFocus()
    If HLen(txtPTName) > 20 Then
        MsgBox "공정명 글자의 자릿수가 너무 많습니다", vbCritical, "공정명 자리수 초과"
        txtPTName.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtPTNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub vsPatternEdit_AfterEdit(ByVal Row As Long, ByVal Col As Long)

On Error GoTo err_msg

    With vsPatternEdit
        If Col = 1 Then
            If .Row = 1 Then
                If Trim(.TextMatrix(Row, 1)) = "" Then
                    .TextMatrix(Row, 1) = "1"
                End If
            Else
                If Trim(.TextMatrix(Row, 1)) = "" Then
                    .TextMatrix(Row, 1) = .TextMatrix(Row - 1, 1)
                End If
            End If
            Call Seq_Sort
        End If
        If Col = 3 Then
            If CInt(.TextMatrix(Row, Col)) >= 0 And CInt(.TextMatrix(Row, Col)) <= 69 Then
                Select Case .TextMatrix(Row, Col)
                    Case "13":  lblMsg = "1)초기, 2)수세1, 3)수세2, 4)염색"
                    Case "14":  lblMsg = "1)초기, 2)수세1, 3)수세2, 4)염색"
                    Case "15":  lblMsg = "1)초기, 2)수세1, 3)수세2, 4)염색"
                    Case "17":  lblMsg = "1)수량, 2)시간, 3)온도"
                    Case "18":  lblMsg = "1)수량, 2)시간, 3)온도"
                    Case "19":  lblMsg = "1)횟수, 2)시간, 3)온도"
                    Case "40":  lblMsg = "1)초기, 2)수세1, 3)수세2, 4)염색"
                    Case "41":  lblMsg = "1,2,3)급수, 4,5,6)리턴"
                    Case "42":  lblMsg = "1,2,3)급수, 4,5,6)리턴"
                    Case Else:  lblMsg = ""
                End Select
                .TextMatrix(Row, Col - 1) = .TextMatrix(Row, Col)
                .TextMatrix(Row, Col) = vsMode.TextMatrix(CInt(.TextMatrix(Row, Col)) + 1, 1)
                
                .Cell(flexcpText, Row, Col + 1, Row, .Cols - 1) = ""
                vsMode.Row = CInt(.TextMatrix(Row, Col - 1)) + 1
                
                If .Row = 1 Then
                    If Trim(.TextMatrix(Row, 1)) = "" Then
                        .TextMatrix(Row, 1) = "1"
                    End If
                Else
                    If Trim(.TextMatrix(Row, 1)) = "" Then
                        .TextMatrix(Row, 1) = .TextMatrix(Row - 1, 1)
                    End If
                End If
                
                If Trim(vsMode.TextMatrix(vsMode.Row, 2)) = "" Then
                    If Trim(vsMode.TextMatrix(vsMode.Row, 6)) = "" Then
                        If .Row = .Rows - 1 Then
                            .Rows = .Rows + 1
                        End If
                        SendKeys "{HOME}"
                        SendKeys "{DOWN}"
                        Call SettingInterval
                        Call GraphRefresh
                    Else
                        SendKeys "{RIGHT}"
                        SendKeys "{RIGHT}"
                        SendKeys "{RIGHT}"
                        .TextMatrix(Row, 9) = vsMode.TextMatrix(vsMode.Row, 9)
                    End If
                Else
                    SendKeys "{RIGHT}"
                    .TextMatrix(Row, 6) = vsMode.TextMatrix(vsMode.Row, 5)
                End If
               
                If vsMode.Row < 11 Then
                    vsMode.TopRow = 0
                ElseIf vsMode.Row < 21 Then
                    vsMode.TopRow = 11
                ElseIf vsMode.Row < 31 Then
                    vsMode.TopRow = 21
                ElseIf vsMode.Row < 41 Then
                    vsMode.TopRow = 31
                ElseIf vsMode.Row < 51 Then
                    vsMode.TopRow = 41
                ElseIf vsMode.Row < 61 Then
                    vsMode.TopRow = 51
                Else
                    vsMode.TopRow = 61
                End If
            End If
        End If
        If Col = 5 Then
            If CInt(.TextMatrix(Row, 2)) >= 0 And CInt(.TextMatrix(Row, 2)) <= 69 Then
                .Cell(flexcpText, Row, Col + 2, Row, .Cols - 1) = ""
                .TextMatrix(Row, Col - 1) = .TextMatrix(Row, Col)
                vsMode.Row = CInt(.TextMatrix(Row, Col - 3)) + 1
            
                If Trim(.TextMatrix(Row, Col)) <> "" Then
                    If CInt(.TextMatrix(Row, Col)) < CInt("0" & vsMode.TextMatrix(vsMode.Row, 3)) Or _
                        CInt(.TextMatrix(Row, Col)) > CInt("0" & vsMode.TextMatrix(vsMode.Row, 4)) Then
                        MsgBox "입력값이 크거나 작습니다. 확인바랍니다", vbCritical, "입력값 오류"
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                Else
                    If Trim(vsMode.TextMatrix(vsMode.Row, 2)) <> "" Then
                        MsgBox "입력값이 있어야합니다", vbCritical, "입력값 오류"
                        Exit Sub
                    End If
                End If
                If Trim(vsMode.TextMatrix(vsMode.Row, 6)) = "" Then
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    SendKeys "{HOME}"
                    SendKeys "{DOWN}"
                    Call SettingInterval
                    Call GraphRefresh
                Else
                    SendKeys "{RIGHT}"
                    SendKeys "{RIGHT}"
                    .TextMatrix(Row, 9) = vsMode.TextMatrix(vsMode.Row, 9)
                End If
                Select Case .TextMatrix(Row, 2)
                    Case "13", "14", "15":
                        Select Case .TextMatrix(Row, Col)
                            Case "1":   .TextMatrix(Row, Col) = "초기"
                            Case "2":   .TextMatrix(Row, Col) = "수세1"
                            Case "3":   .TextMatrix(Row, Col) = "수세2"
                            Case "4":   .TextMatrix(Row, Col) = "염색"
                        End Select
                    Case "17":
                        Select Case .TextMatrix(Row, Col)
                            Case "1":   .TextMatrix(Row, Col + 4) = "L"
                                        .TextMatrix(Row, Col) = "수량"
                                        lblMsg = "1 ~ 999 (L)"
                            Case "2":   .TextMatrix(Row, Col + 4) = "분"
                                        .TextMatrix(Row, Col) = "시간"
                                        lblMsg = "1 ~ 99 (분)"
                            Case "3":   .TextMatrix(Row, Col + 4) = "도"
                                        .TextMatrix(Row, Col) = "온도"
                                        lblMsg = "1 ~ 99 (도)"
                            Case Else:  lblMsg = ""
                        End Select
                    Case "18":
                        Select Case .TextMatrix(Row, Col)
                            Case "1":   .TextMatrix(Row, Col + 4) = "L"
                                        .TextMatrix(Row, Col) = "수량"
                                        lblMsg = "1 ~ 999 (L)"
                            Case "2":   .TextMatrix(Row, Col + 4) = "분"
                                        .TextMatrix(Row, Col) = "시간"
                                        lblMsg = "1 ~ 99 (분)"
                            Case "3":   .TextMatrix(Row, Col + 4) = "도"
                                        .TextMatrix(Row, Col) = "온도"
                                        lblMsg = "1 ~ 99 (도)"
                            Case Else:  lblMsg = ""
                        End Select
                    Case "19":
                        Select Case .TextMatrix(Row, Col)
                            Case "1":   .TextMatrix(Row, Col + 4) = "회"
                                        .TextMatrix(Row, Col) = "회수"
                                        lblMsg = "1 ~ 99 (회)"
                            Case "2":   .TextMatrix(Row, Col + 4) = "분"
                                        .TextMatrix(Row, Col) = "시간"
                                        lblMsg = "1 ~ 99 (분)"
                            Case "3":   .TextMatrix(Row, Col + 4) = "도"
                                        .TextMatrix(Row, Col) = "온도"
                                        lblMsg = "1 ~ 99 (도)"
                            Case Else:  lblMsg = ""
                        End Select
                    Case "40":
                        Select Case .TextMatrix(Row, Col)
                            Case "1":   .TextMatrix(Row, Col + 4) = "도"
                                        .TextMatrix(Row, Col) = "초기"
                                        lblMsg = "1 ~ 99 (도)"
                            Case "2":   .TextMatrix(Row, Col + 4) = ""
                                        .TextMatrix(Row, Col) = "수세1"
                                        lblMsg = ""
                            Case "3":   .TextMatrix(Row, Col + 4) = ""
                                        .TextMatrix(Row, Col) = "수세2"
                                        lblMsg = ""
                            Case "4":   .TextMatrix(Row, Col + 4) = ""
                                        .TextMatrix(Row, Col) = "염색"
                                        lblMsg = ""
                            Case Else:  lblMsg = ""
                        End Select
                    Case "41", "42":
                        Select Case .TextMatrix(Row, Col)
                            Case "1", "2", "3": .TextMatrix(Row, Col + 4) = "도"
                                                .TextMatrix(Row, Col) = "급수"
                                                lblMsg = "1 ~ 99 (도)"
                            Case "4", "5", "6": .TextMatrix(Row, Col + 4) = ""
                                                .TextMatrix(Row, Col) = "리턴"
                                                lblMsg = ""
                            Case Else:  lblMsg = ""
                        End Select
                    Case Else:  lblMsg = ""
                End Select
            End If
        End If
        If Col = 8 Then
            If CInt(.TextMatrix(Row, 2)) >= 0 And CInt(.TextMatrix(Row, 2)) <= 69 Then
                vsMode.Row = CInt(.TextMatrix(Row, Col - 6)) + 1
                If Trim(.TextMatrix(Row, Col)) <> "" Then
                    If CInt(.TextMatrix(Row, Col)) < CInt(vsMode.TextMatrix(vsMode.Row, 7)) Or _
                        CInt(.TextMatrix(Row, Col)) > CInt(vsMode.TextMatrix(vsMode.Row, 8)) Then
                        lblMsg = ""
                        MsgBox "입력값이 크거나 작습니다. 확인바랍니다", vbCritical, "입력값 오류"
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                Else
                    If Trim(vsMode.TextMatrix(vsMode.Row, 6)) <> "" Then
                        lblMsg = ""
                        MsgBox "입력값이 있어야합니다", vbCritical, "입력값 오류"
                        Exit Sub
                    End If
                End If
                .TextMatrix(Row, Col - 1) = .TextMatrix(Row, Col)
                lblMsg = ""
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                End If
                SendKeys "{HOME}"
                SendKeys "{DOWN}"
                Call SettingInterval
                Call GraphRefresh
            End If
        End If
    End With
    
    Exit Sub
    
err_msg:
    MsgBox "입력값이 잘못되었습니다" & vbCrLf & "확인후 재작업 해주십시요", vbCritical, "입력값 오류"
End Sub

Private Sub vsPatternEdit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPatternEdit
        If Col = 6 Or Col = 9 Then
            Cancel = True
        Else
            Cancel = False
        End If
    End With
End Sub

Private Sub vsPatternEdit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
Dim iCount As Integer

    With vsPatternEdit

        .Redraw = flexRDBuffered
        If KeyAscii = vbKeyReturn Then
            Select Case Col
                Case 6
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    SendKeys "{HOME}"
                    SendKeys "{DOWN}"

                Case 1
                    SendKeys "{RIGHT}"
'                Case 3
'                    SendKeys "{RIGHT}"
            End Select
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsPatternList_RowColChange()
    Call vsPatternList_Click
End Sub

Private Sub txtPTNo_LostFocus()
Dim oDyeCon As PlusLib2.CDyeCon
Dim tRs As ADODB.Recordset
Dim i%

    On Error GoTo ErrHandler

    Set oDyeCon = New PlusLib2.CDyeCon
    oDyeCon.Connection = g_adoCon
    oDyeCon.UserName = g_sUserName

    Set tRs = oDyeCon.GetPattern(1, 0, CInt(txtPTNo))
    If tRs.RecordCount > 0 Then
        txtPTName = tRs!PtName
        Call vsPatternEditSet
        With vsPatternEdit
            .Rows = tRs.RecordCount + 1
            For i = 1 To tRs.RecordCount
                .TextMatrix(i, 1) = tRs!worksection
                .TextMatrix(i, 2) = tRs!ModeNo
                .TextMatrix(i, 3) = tRs!ModeName
                If tRs!SelNo1 = 0 Then
                    .TextMatrix(i, 4) = ""
                    .TextMatrix(i, 5) = ""
                Else
                    .TextMatrix(i, 4) = tRs!SelNo1
                    .TextMatrix(i, 5) = tRs!selectunit1
                End If
                If tRs!SelNo2 = 0 Then
                    .TextMatrix(i, 6) = ""
                    .TextMatrix(i, 7) = tRs!selectunit2
                Else
                    .TextMatrix(i, 6) = tRs!SelNo2
                    .TextMatrix(i, 7) = tRs!selectunit2
                End If
                tRs.MoveNext
            Next i
        End With
    Else
        Call vsPatternEditSet
        txtPTName = ""
    End If
    tRs.Close
    Set tRs = Nothing
    Set oDyeCon = Nothing
    txtPTName.SetFocus
    Screen.MousePointer = vbDefault
    
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    Set tRs = Nothing
    Set oDyeCon = Nothing
    Call ErrorBox(Err.Number, "frmDyePattern.txtPTNo_LostFocus", Err.Description)
End Sub

Private Sub cmdKill_Click()
    Dim oDyeCon As PlusLib2.CDyeCon
    
    If Trim(txtPTNo) = "" Then
        MsgBox "삭제할 염색패턴이 선택되지 않습니다" & vbCrLf & "확인후 재작업하여주십시요", vbCritical, "오류"
        Exit Sub
    End If
    If MsgBox("정말로 삭제하겠습니까?", vbYesNo, "선택") = vbYes Then
        Set oDyeCon = New PlusLib2.CDyeCon
        oDyeCon.Connection = g_adoCon
        oDyeCon.UserName = g_sUserName
        g_adoCon.BeginTrans
        If oDyeCon.DeletePattern(1, 0, CInt(txtPTNo)) Then
            g_adoCon.CommitTrans
            MsgBox "성공적으로 삭제되었습니다", vbOKOnly, "삭제 성공"
            Call vsPatternSet
            Call vsPatternFill
            txtPTNo = ""
            txtPTName = ""
            Call vsPatternEditSet
        End If
        Set oDyeCon = Nothing
    End If
End Sub

Private Sub cmdSave_Click()
Dim oDyeCon As PlusLib2.CDyeCon
Dim tPattern As PlusLib2.TWorkPattern
Dim tRs As ADODB.Recordset
Dim iRow As Integer
Dim iSection As Integer
Dim iSectionSeq As Integer
Dim bSucc As Boolean

    If Trim(txtPTNo) = "" Or Trim(txtPTName) = "" Then
        MsgBox "염색패턴번호나 염색패턴명이 잘못되어있습니다" & vbCrLf & "확인후 재작업하여주십시요", vbCritical, "오류"
        Exit Sub
    End If
    With vsPatternEdit
        If .Rows > 2 Then
            Set oDyeCon = New PlusLib2.CDyeCon
            oDyeCon.Connection = g_adoCon
            oDyeCon.UserName = g_sUserName
            
'            Set tRs = oDyeCon.GetPatternGroup(1, 0, CInt(txtPTNo))
            bSucc = True
            g_adoCon.BeginTrans
'            If tRs.RecordCount > 0 Then
                If oDyeCon.DeletePattern(1, 0, CInt(txtPTNo)) Then
                    For iRow = 1 To .Rows - 1
                        If Trim(.TextMatrix(iRow, 2)) <> "" Then
                            If Trim(.TextMatrix(iRow, 1)) = "" Then
                                iSectionSeq = iSectionSeq + 1
                            Else
                                If iSection = CInt(.TextMatrix(iRow, 1)) Then
                                    iSectionSeq = iSectionSeq + 1
                                Else
                                    iSectionSeq = 1
                                End If
                                iSection = CInt(.TextMatrix(iRow, 1))
                            End If
                            tPattern.DyeKind = 1
                            tPattern.DyeID = 0
                            tPattern.PtNo = CInt(txtPTNo)
                            tPattern.PtName = Trim(txtPTName)
                            tPattern.Section = iSection
                            tPattern.Seq = CInt(iSectionSeq)
                            tPattern.ModeNo = CInt(.TextMatrix(iRow, 2))
                            tPattern.SelNo1 = CInt("0" & .TextMatrix(iRow, 4))
                            tPattern.SelNo2 = CInt("0" & .TextMatrix(iRow, 7))
                            
                            If Not oDyeCon.InsertPattern(tPattern) Then
                                bSucc = False
                                Exit For
                            End If
                        End If
                    Next iRow
                Else
                    bSucc = False
                End If
'            End If
'            tRs.Close
'            Set tRs = Nothing
            If bSucc = True Then
                g_adoCon.CommitTrans
                MsgBox "성공적으로 저장되었습니다", vbOKOnly, "저장 성공"
                Call vsPatternSet
                Call vsPatternFill
            End If
            Set oDyeCon = Nothing
        End If
    End With
End Sub
















