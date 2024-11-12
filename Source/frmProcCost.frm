VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProcCost 
   Caption         =   "카드별 진행"
   ClientHeight    =   9450
   ClientLeft      =   75
   ClientTop       =   525
   ClientWidth     =   15420
   Icon            =   "frmProcCost.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   15420
   Begin VB.Frame fraAdd 
      Caption         =   "[[ 가공료 데이터 추가 ]]"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   270
      TabIndex        =   29
      Top             =   7950
      Width           =   9345
      Begin VB.ComboBox cboData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   4230
         Style           =   2  '드롭다운 목록
         TabIndex        =   46
         Top             =   1710
         Width           =   915
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   3960
         TabIndex        =   78
         Top             =   3150
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ComboBox cboData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   4230
         Style           =   2  '드롭다운 목록
         TabIndex        =   54
         Top             =   1320
         Width           =   915
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   2760
         TabIndex        =   52
         Top             =   3150
         Width           =   1065
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   6600
         TabIndex        =   57
         Top             =   2445
         Width           =   2505
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   6600
         TabIndex        =   58
         Top             =   2805
         Width           =   2505
      End
      Begin VB.ComboBox cboData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   3090
         Style           =   2  '드롭다운 목록
         TabIndex        =   49
         Top             =   2415
         Width           =   765
      End
      Begin Threed.SSPanel pnlYYMM 
         Height          =   315
         Left            =   1620
         TabIndex        =   64
         Top             =   420
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   556
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.ComboBox cboData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   2670
         Style           =   2  '드롭다운 목록
         TabIndex        =   61
         Top             =   4455
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.ComboBox cboData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6600
         Style           =   2  '드롭다운 목록
         TabIndex        =   55
         Top             =   1710
         Width           =   2505
      End
      Begin VB.ComboBox cboData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   6600
         Style           =   2  '드롭다운 목록
         TabIndex        =   53
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox cboData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1620
         Style           =   2  '드롭다운 목록
         TabIndex        =   50
         Top             =   2775
         Width           =   2235
      End
      Begin VB.ComboBox cboData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1620
         Style           =   2  '드롭다운 목록
         TabIndex        =   47
         Top             =   2055
         Width           =   2235
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   44
         Top             =   1320
         Width           =   2235
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1620
         TabIndex        =   45
         Top             =   1695
         Width           =   2235
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1620
         TabIndex        =   48
         Top             =   2415
         Width           =   1455
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1620
         TabIndex        =   51
         Top             =   3150
         Width           =   1065
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   6600
         TabIndex        =   56
         Top             =   2085
         Width           =   2505
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   2670
         TabIndex        =   62
         Top             =   4830
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   2670
         TabIndex        =   63
         Top             =   5190
         Visible         =   0   'False
         Width           =   2505
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   20
         Left            =   270
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Order No"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   21
         Left            =   270
         TabIndex        =   31
         Top             =   1695
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "품      명"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   22
         Left            =   270
         TabIndex        =   32
         Top             =   2055
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "가공 구분"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   23
         Left            =   270
         TabIndex        =   33
         Top             =   2775
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "정산 구분"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   25
         Left            =   270
         TabIndex        =   34
         Top             =   2415
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "가공 단가"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   27
         Left            =   270
         TabIndex        =   35
         Top             =   3150
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "계상 수량"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   28
         Left            =   5250
         TabIndex        =   36
         Top             =   1710
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "부가세구분"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   29
         Left            =   5250
         TabIndex        =   37
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "거래 구분"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   30
         Left            =   1320
         TabIndex        =   38
         Top             =   4455
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "화폐 단위"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   31
         Left            =   1320
         TabIndex        =   39
         Top             =   4830
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "외화 환율"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   32
         Left            =   1320
         TabIndex        =   40
         Top             =   5190
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "외화 금액"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   33
         Left            =   5250
         TabIndex        =   41
         Top             =   2085
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "청구 금액"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   24
         Left            =   270
         TabIndex        =   42
         Top             =   780
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "거래처"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   34
         Left            =   270
         TabIndex        =   43
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "적용 년월"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlCustom 
         Height          =   315
         Left            =   1620
         TabIndex        =   65
         Top             =   780
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   556
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   0
         Left            =   3870
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1320
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
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   3870
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1710
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
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   35
         Left            =   5250
         TabIndex        =   69
         Top             =   2805
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "합계 금액"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel pnlName 
         Height          =   315
         Index           =   36
         Left            =   5250
         TabIndex        =   70
         Top             =   2445
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "부  가  세"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdAddSave 
         Height          =   630
         Left            =   5400
         TabIndex        =   59
         Top             =   3660
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   1111
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "입력저장"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdAddCancel 
         Height          =   630
         Left            =   7350
         TabIndex        =   60
         Top             =   3660
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   1111
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "입력취소"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   150
         X2              =   9240
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         Height          =   3285
         Left            =   150
         Top             =   300
         Width           =   9105
      End
   End
   Begin Threed.SSCommand cmdComplete 
      Height          =   630
      Left            =   3750
      TabIndex        =   68
      Top             =   8550
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1111
      _Version        =   196609
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "계상완료"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Frame fraDetail 
      Caption         =   "[[ 계상처리 입력사항 ]]"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3030
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   2745
      Begin VB.Frame fraInputDate 
         Caption         =   "기간 선택"
         Height          =   585
         Left            =   30
         TabIndex        =   16
         Top             =   1080
         Width           =   2685
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   0
            Left            =   30
            TabIndex        =   17
            Top             =   240
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            Format          =   120258561
            CurrentDate     =   36871
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Index           =   1
            Left            =   1410
            TabIndex        =   18
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            Format          =   120258561
            CurrentDate     =   36871
         End
         Begin VB.Label Label3 
            Caption         =   "∼"
            Height          =   165
            Left            =   1260
            TabIndex        =   22
            Top             =   300
            Width           =   225
         End
      End
      Begin VB.Frame fraInputItem 
         Caption         =   "정산료 구분"
         Height          =   585
         Left            =   30
         TabIndex        =   13
         Top             =   300
         Width           =   2685
         Begin VB.OptionButton optAccount 
            Caption         =   "출고량정산"
            Height          =   180
            Index           =   1
            Left            =   1380
            TabIndex        =   15
            Top             =   330
            Width           =   1215
         End
         Begin VB.OptionButton optAccount 
            Caption         =   "오더량정산"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   330
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   540
         Left            =   1410
         TabIndex        =   21
         Tag             =   "PERM_ADDNEW"
         Top             =   1740
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   953
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
         Caption         =   "작성 취소"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   540
         Left            =   60
         TabIndex        =   20
         Tag             =   "PERM_ADDNEW"
         Top             =   1740
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   953
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
         Caption         =   "작성 완료"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin VB.Frame fraInputExchRate 
         Caption         =   "적용 환율( $ 환율 )"
         Height          =   585
         Left            =   30
         TabIndex        =   27
         Top             =   1860
         Visible         =   0   'False
         Width           =   2685
         Begin VB.TextBox txtExchRate 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1260
            TabIndex        =   19
            Top             =   240
            Width           =   1365
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   300
            Left            =   30
            TabIndex        =   28
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   -2147483644
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "미국 USD"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
      End
   End
   Begin Threed.SSPanel pnlSumUp 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   16325
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Frame fraSearch 
         Height          =   945
         Left            =   30
         TabIndex        =   2
         Top             =   -60
         Width           =   5850
         Begin Threed.SSPanel SSPanel1 
            Height          =   345
            Left            =   60
            TabIndex        =   83
            Top             =   180
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
            _Version        =   196609
            Caption         =   "계상년월"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.TextBox txtSearch 
            Height          =   300
            Index           =   1
            Left            =   1290
            TabIndex        =   71
            Top             =   570
            Width           =   1635
         End
         Begin VB.ComboBox cboMonth 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2475
            Sorted          =   -1  'True
            TabIndex        =   4
            Text            =   "01"
            Top             =   210
            Width           =   615
         End
         Begin VB.ComboBox cboYear 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   3
            Text            =   "2002"
            Top             =   210
            Width           =   885
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   690
            Left            =   3390
            TabIndex        =   5
            Tag             =   "PERM_ADDNEW"
            Top             =   150
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   1217
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
            Caption         =   "         조회"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSCommand cmdSumUp 
            Height          =   690
            Left            =   4590
            TabIndex        =   8
            Tag             =   "PERM_ADDNEW"
            Top             =   150
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   1217
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
            Caption         =   "       계상처리"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   60
            TabIndex        =   72
            Top             =   570
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   196609
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkSearch 
               Caption         =   "거 래 처"
               Height          =   240
               Index           =   1
               Left            =   60
               TabIndex        =   73
               Top             =   45
               Width           =   975
            End
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   2
            Left            =   2940
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   570
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "월"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   3090
            TabIndex        =   7
            Top             =   270
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "년"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   2190
            TabIndex        =   6
            Top             =   270
            Width           =   255
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdOut 
         Height          =   7095
         Left            =   30
         TabIndex        =   1
         Top             =   1410
         Width           =   5790
         _cx             =   10213
         _cy             =   12515
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   585
         Left            =   60
         TabIndex        =   80
         Top             =   8550
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   1032
         _Version        =   196609
         CaptionStyle    =   1
         BackColor       =   12648447
         Caption         =   "※ 선택한 거래처에 한하여 계상완료 처리 합니다"
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdCheck 
         Height          =   480
         Index           =   0
         Left            =   30
         TabIndex        =   81
         Tag             =   "PERM_ADDNEW"
         Top             =   900
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   847
         _Version        =   196609
         CaptionStyle    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전체선택"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdCheck 
         Height          =   480
         Index           =   1
         Left            =   1200
         TabIndex        =   82
         Tag             =   "PERM_ADDNEW"
         Top             =   900
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   847
         _Version        =   196609
         CaptionStyle    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "선택해제"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel pnlProcCost 
      Height          =   9255
      Left            =   5850
      TabIndex        =   9
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   16325
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboName 
         Height          =   300
         Index           =   22
         Left            =   5400
         Style           =   2  '드롭다운 목록
         TabIndex        =   147
         Top             =   570
         Width           =   1155
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   1545
         Left            =   30
         TabIndex        =   132
         Top             =   7680
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   2725
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtAddTaxSeq 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   1470
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   135
            Top             =   60
            Width           =   525
         End
         Begin VB.TextBox txtAddTaxSeq 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   2130
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   134
            Top             =   60
            Width           =   525
         End
         Begin VB.TextBox txtAddTaxSeq 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   2790
            MaxLength       =   4
            TabIndex        =   133
            Top             =   60
            Width           =   855
         End
         Begin Threed.SSPanel pnlName 
            Height          =   345
            Index           =   6
            Left            =   60
            TabIndex        =   136
            Top             =   60
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "일련 No"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   345
            Index           =   7
            Left            =   60
            TabIndex        =   137
            Top             =   420
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "공급가 액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   345
            Index           =   8
            Left            =   60
            TabIndex        =   138
            Top             =   780
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "부가 세"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlSumSupplyPrice 
            Height          =   345
            Index           =   0
            Left            =   1470
            TabIndex        =   139
            Top             =   420
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlAddTaxPrice 
            Height          =   345
            Left            =   1470
            TabIndex        =   140
            Top             =   780
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   345
            Index           =   5
            Left            =   60
            TabIndex        =   141
            Top             =   1140
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "합계금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlAddTaxSum 
            Height          =   345
            Left            =   1470
            TabIndex        =   142
            Top             =   1140
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlMaxTaxSeq 
            Height          =   345
            Left            =   3690
            TabIndex        =   143
            Top             =   450
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   12648384
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpPrnDate 
            Height          =   375
            Left            =   5100
            TabIndex        =   144
            Top             =   60
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   120258560
            CurrentDate     =   38125
         End
         Begin Threed.SSPanel pnlName 
            Height          =   345
            Index           =   10
            Left            =   3690
            TabIndex        =   145
            Top             =   60
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "청구 일자"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   660
            Left            =   5340
            TabIndex        =   146
            Top             =   840
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   1164
            _Version        =   196609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "    계산서 발행"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            Index           =   1
            X1              =   2010
            X2              =   2160
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            Index           =   2
            X1              =   2670
            X2              =   2820
            Y1              =   240
            Y2              =   240
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   600
         Left            =   30
         TabIndex        =   124
         Top             =   4620
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   1058
         _Version        =   196609
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel pnlCount 
            Height          =   315
            Left            =   1260
            TabIndex        =   125
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlQty 
            Height          =   315
            Left            =   2370
            TabIndex        =   127
            Top             =   120
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlSumPrice 
            Height          =   315
            Left            =   4050
            TabIndex        =   129
            Top             =   120
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Alignment       =   4
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   15
            Left            =   30
            TabIndex        =   131
            Top             =   120
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   196609
            CaptionStyle    =   1
            ForeColor       =   16777215
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "계"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "원"
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
            Index           =   3
            Left            =   5490
            TabIndex        =   130
            Top             =   180
            Width           =   210
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "YDS"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   3510
            TabIndex        =   128
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "건"
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
            Index           =   1
            Left            =   2040
            TabIndex        =   126
            Top             =   180
            Width           =   210
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   2445
         Left            =   30
         TabIndex        =   84
         Top             =   5250
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   4313
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkOrderFlag 
            BackColor       =   &H00FFC0C0&
            Caption         =   "사용"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   6390
            TabIndex        =   149
            Top             =   780
            Width           =   1815
         End
         Begin VB.ComboBox cboCurrency 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10560
            Style           =   2  '드롭다운 목록
            TabIndex        =   123
            Top             =   90
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtUnitPrice 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1350
            TabIndex        =   108
            Top             =   750
            Width           =   1875
         End
         Begin VB.TextBox txtSumQty 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2310
            TabIndex        =   107
            Top             =   1080
            Width           =   915
         End
         Begin VB.ComboBox cboDealClss 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4590
            Style           =   2  '드롭다운 목록
            TabIndex        =   106
            Top             =   750
            Width           =   1755
         End
         Begin VB.TextBox txtExchangeRate 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10560
            TabIndex        =   105
            Top             =   435
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.TextBox txtSupplyPrice 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1350
            TabIndex        =   104
            Top             =   2085
            Width           =   1875
         End
         Begin VB.TextBox txtForeignPrice 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10560
            TabIndex        =   103
            Top             =   765
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.ComboBox cboAdjustClss 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4590
            Style           =   2  '드롭다운 목록
            TabIndex        =   102
            Top             =   1110
            Width           =   1755
         End
         Begin VB.TextBox txtTax 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1350
            TabIndex        =   101
            Top             =   1740
            Width           =   1875
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1350
            TabIndex        =   100
            Top             =   1410
            Width           =   1875
         End
         Begin VB.TextBox txtOutQty 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1350
            TabIndex        =   99
            Top             =   1080
            Width           =   915
         End
         Begin VB.ComboBox cboOrderFlag 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4590
            Style           =   2  '드롭다운 목록
            TabIndex        =   98
            Top             =   1470
            Width           =   1755
         End
         Begin VB.TextBox txtPrevOutQty 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7710
            Locked          =   -1  'True
            TabIndex        =   91
            Top             =   60
            Width           =   1545
         End
         Begin VB.TextBox txtOrderQty 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4590
            Locked          =   -1  'True
            TabIndex        =   90
            Top             =   375
            Width           =   1365
         End
         Begin VB.TextBox txtWorkName 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4590
            Locked          =   -1  'True
            TabIndex        =   89
            Top             =   45
            Width           =   1725
         End
         Begin VB.TextBox txtArticle 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   360
            Width           =   1875
         End
         Begin VB.TextBox txtOrderNo 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   30
            Width           =   1875
         End
         Begin VB.TextBox txtUnitClss 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5970
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   375
            Width           =   345
         End
         Begin VB.TextBox txtPrevSumQty 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7710
            Locked          =   -1  'True
            TabIndex        =   85
            Top             =   390
            Width           =   1545
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   0
            Left            =   30
            TabIndex        =   92
            Top             =   30
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Order No."
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   1
            Left            =   30
            TabIndex        =   93
            Top             =   360
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "품       명"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   2
            Left            =   3270
            TabIndex        =   94
            Top             =   45
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "가공  구분"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   3
            Left            =   3270
            TabIndex        =   95
            Top             =   375
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "수 주  량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   4
            Left            =   6360
            TabIndex        =   96
            Top             =   60
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "전월누계출고"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   38
            Left            =   6360
            TabIndex        =   97
            Top             =   390
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "전월누계청구"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   12
            Left            =   30
            TabIndex        =   109
            Top             =   750
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "가공단가"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   13
            Left            =   30
            TabIndex        =   110
            Top             =   1080
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "출고수량"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   14
            Left            =   30
            TabIndex        =   111
            Top             =   2085
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "합계금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   330
            Index           =   16
            Left            =   3270
            TabIndex        =   112
            Top             =   750
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkFreeTax 
               BackColor       =   &H00FFC0C0&
               Caption         =   "영세율"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   90
               TabIndex        =   150
               Top             =   60
               Width           =   1095
            End
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   17
            Left            =   9630
            TabIndex        =   113
            Top             =   90
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "화폐단위"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   18
            Left            =   9630
            TabIndex        =   114
            Top             =   435
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "환율"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   300
            Index           =   19
            Left            =   9630
            TabIndex        =   115
            Top             =   765
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "외화금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdUpdate 
            Height          =   540
            Left            =   7020
            TabIndex        =   116
            Top             =   1830
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   196609
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "저장"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   540
            Left            =   5850
            TabIndex        =   117
            Top             =   1830
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   196609
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "추가"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlName 
            Height          =   330
            Index           =   26
            Left            =   3270
            TabIndex        =   118
            Top             =   1110
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "정산구분"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   540
            Left            =   8190
            TabIndex        =   119
            Top             =   1830
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   953
            _Version        =   196609
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "삭제"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   9
            Left            =   30
            TabIndex        =   120
            Top             =   1740
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "부  가  세"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   315
            Index           =   37
            Left            =   30
            TabIndex        =   121
            Top             =   1410
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "청구 금액"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlName 
            Height          =   330
            Index           =   39
            Left            =   3270
            TabIndex        =   122
            Top             =   1470
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "오더구분"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   600
         Left            =   5820
         TabIndex        =   26
         Top             =   4620
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1058
         _Version        =   196609
         CaptionStyle    =   1
         BackColor       =   12648447
         Caption         =   "※ 선택한 오더에 한하여     수주 마감처리합니다"
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   810
         Left            =   7710
         TabIndex        =   10
         Tag             =   "PERM_ADDNEW"
         Top             =   8340
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1429
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
         Caption         =   "       닫기"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdCloseAndComp 
         Height          =   600
         Left            =   8010
         TabIndex        =   11
         Top             =   4620
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   1058
         _Version        =   196609
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "수주마감"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel pnlTerm 
         Height          =   315
         Left            =   60
         TabIndex        =   23
         Top             =   90
         Visible         =   0   'False
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         AutoSize        =   1
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdCheck 
         Height          =   480
         Index           =   2
         Left            =   30
         TabIndex        =   24
         Tag             =   "PERM_ADDNEW"
         Top             =   450
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   847
         _Version        =   196609
         CaptionStyle    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전체선택"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdCheck 
         Height          =   480
         Index           =   3
         Left            =   1260
         TabIndex        =   25
         Tag             =   "PERM_ADDNEW"
         Top             =   450
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   847
         _Version        =   196609
         CaptionStyle    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "선택해제"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSFrame fraOrder 
         Height          =   405
         Left            =   4380
         TabIndex        =   75
         Top             =   90
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   714
         _Version        =   196609
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   120
            Value           =   -1  'True
            Width           =   1260
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리 번호"
            Height          =   180
            Index           =   1
            Left            =   1800
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   120
            Width           =   1290
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdSumUp 
         Height          =   3645
         Left            =   30
         TabIndex        =   79
         Top             =   930
         Width           =   9300
         _cx             =   16404
         _cy             =   6429
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   315
         Left            =   4380
         TabIndex        =   148
         Top             =   570
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "사용구분"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   60
      Top             =   9300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmProcCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'
'변경이력
'
'2013.12.12   자체    오승욱   S_201312_태을염직_99   지번주소에서 도로명 주소로 입력가능하게,거래처 주소 도로명 주소 Select
'**************************************************************************************************
Option Explicit
Private ThisMon As String
Private LastMon As String
'Private Prev2Mon As String

Private m_CustomID As String
Private m_sOrderID As String
Private m_nWorkUnitPrice As Single
Private m_sPrinter As String
Private m_bloading As Boolean
Private m_bLoading1  As Boolean
Private m_bLoading2  As Boolean

Private Sub cboAdjustClss_Click()
    Dim i%
    If m_bLoading2 Then Exit Sub
    
    If cboAdjustClss.Text = "오더" Then
        If CheckNum(txtOrderQty) - CheckNum(txtPrevOutQty) > grdSumUp.TextMatrix(grdSumUp.Row, 25) Then
            txtOutQty = grdSumUp.TextMatrix(grdSumUp.Row, 25)
        Else
            txtOutQty = Format(CheckNum(txtOrderQty) - CheckNum(txtPrevOutQty), "#,##0")
        End If
    Else
        txtOutQty = grdSumUp.TextMatrix(grdSumUp.Row, 25)
    End If
    
    If txtUnitClss = "Y" Then
        txtSumQty = txtOutQty
    Else
        txtSumQty = Format(CheckNum(txtOutQty) * 1.0936, "#,##0")
    End If
    
    
    If cboCurrency.Text = "$" Then
        If Trim(txtExchangeRate.Text) <> "" And Trim(txtExchangeRate.Text) <> "0" Then
            txtAmount = Format(CheckNum(txtSumQty) * CSng(txtUnitPrice) * CheckNum(txtExchangeRate), "#,##0")
            txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "#,##0")
            txtSupplyPrice = Format(CheckNum(txtAmount) + CheckNum(txtTax), "#,##0")
        
            txtForeignPrice = Format(CheckNum(txtAmount) / CSng(txtExchangeRate), "$#,##0.00")
        Else
            txtAmount = Format(CheckNum(txtSumQty) * CSng(txtUnitPrice) * CheckNum(grdSumUp.TextMatrix(grdSumUp.Row, 15)), "#,##0")
            txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "#,##0")
            txtSupplyPrice = Format(CheckNum(txtAmount) + CheckNum(txtTax), "#,##0")
        
            txtExchangeRate = grdSumUp.TextMatrix(grdSumUp.Row, 15)
            txtForeignPrice = Format(CheckNum(txtAmount) / CSng(txtExchangeRate), "$#,##0.00")
        End If
        
        
    Else
        txtExchangeRate = ""
        txtForeignPrice = ""
        
        txtAmount = Format(CheckNum(txtSumQty) * CSng(txtUnitPrice), "#,##0")
        txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "#,##0")
        txtSupplyPrice = Format(CheckNum(txtAmount) + CheckNum(txtTax), "#,##0")
        
    End If
    
End Sub

Private Sub cboCurrency_Click()

On Error GoTo ErrHandler
    If m_bLoading2 Then Exit Sub
    
    If cboCurrency.Text = "$" Then
        If Trim(txtExchangeRate.Text) <> "" And Trim(txtExchangeRate.Text) <> "0" Then
            txtAmount = Format(CheckNum(txtSumQty) * CSng(txtUnitPrice) * CheckNum(txtExchangeRate), "#,##0")
            txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "#,##0")
            txtSupplyPrice = Format(CheckNum(txtAmount) + CheckNum(txtTax), "#,##0")
        
            txtForeignPrice = Format(CheckNum(txtAmount) / CSng(txtExchangeRate), "$#,##0.00")
        Else
            txtAmount = Format(CheckNum(txtSumQty) * CSng(txtUnitPrice) * CheckNum(grdSumUp.TextMatrix(grdSumUp.Row, 15)), "#,##0")
            txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "#,##0")
            txtSupplyPrice = Format(CheckNum(txtAmount) + CheckNum(txtTax), "#,##0")
        
            txtExchangeRate = grdSumUp.TextMatrix(grdSumUp.Row, 15)
            txtForeignPrice = Format(CheckNum(txtAmount) / CSng(txtExchangeRate), "$#,##0.00")
        End If
        
        
    Else
        txtExchangeRate = ""
        txtForeignPrice = ""
        
        txtAmount = Format(CheckNum(txtSumQty) * CSng(txtUnitPrice), "#,##0")
        txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "#,##0")
        txtSupplyPrice = Format(CheckNum(txtAmount) + CheckNum(txtTax), "#,##0")
        
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "수치 입력이 잘못되었습니다", vbExclamation + vbOKOnly, "입력오류"
End Sub

Private Sub cboCurrency_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Call NextFocus
    End If
End Sub

Private Sub cboData_Click(Index As Integer)
    If m_bLoading2 Then Exit Sub
    
    If Index = 3 Or Index = 4 Then
        txtData(4) = Format(CheckNum(txtData(9)) * CSng(SetCurrency(txtData(2), 2)), "##,##0")
        If cboData(3).Text = "포함" Then
            txtData(8) = Format(Fix(CheckNum((txtData(4))) * 0.1), "##,##0")
        Else
            txtData(8) = "0"
        End If
        txtData(7) = Format(CheckNum(txtData(4)) + CheckNum(txtData(8)), "#,##0")
        
        If cboData(4).Text = "$" Then
            If Trim(txtData(5).Text) <> "" And Trim(txtData(5).Text) <> "0" Then
                txtData(6) = Format(CheckNum(txtData(3)) * CheckNum(txtData(2)) / CheckNum(txtData(5)), "$#,##0.00")
            Else
                txtData(6) = ""
            End If
        Else
            txtData(5) = ""
            txtData(6) = ""
        End If
    End If
End Sub

Private Sub cboData_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call NextFocus
    End If
End Sub

Private Sub cboDealClss_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Call NextFocus
    End If
End Sub




Private Sub cboName_Click(Index As Integer)
    If m_bloading = True And Index = 22 Then
        If grdOut.Rows >= grdOut.FixedRows Then
            Call FillGridSumUp
        End If
    End If
End Sub

Private Sub chkFreeTax_Click()
    cboDealClss.Enabled = chkFreeTax.Value
        
    If Trim(txtSumQty) <> "" And Trim(txtUnitPrice) <> "" Then
        If cboCurrency.Text = "$" Then
            If Trim(txtExchangeRate.Text) <> "" And Trim(txtExchangeRate.Text) <> "0" Then
                txtAmount = Format(CheckNum(txtSumQty) * CSng(txtUnitPrice) * CheckNum(txtExchangeRate), "#,##0")
                If chkFreeTax.Value = vbUnchecked Then
                    txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "#,##0")
                Else
                    txtTax = "0"
                End If
                txtSupplyPrice = Format(CheckNum(txtAmount) + CheckNum(txtTax), "#,##0")
                
                If Trim(txtExchangeRate.Text) <> "" And Trim(txtExchangeRate.Text) <> "0" Then
                    txtForeignPrice = Format(CheckNum(txtAmount) / CheckNum(txtExchangeRate), "$#,##0.00")
                Else
                    txtForeignPrice = ""
                End If
            End If
        Else
            txtAmount = Format(CheckNum(txtSumQty) * CSng(txtUnitPrice), "#,##0")
            If chkFreeTax.Value = vbUnchecked Then
                txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "#,##0")
            Else
                txtTax = "0"
            End If
            txtSupplyPrice = Format(CheckNum(txtAmount) + CheckNum(txtTax), "#,##0")
        End If
    End If
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If chkSearch(Index).Value = vbChecked Then
        txtSearch(Index).Enabled = True
        txtSearch(Index).SetFocus
    Else
        txtSearch(Index).Enabled = False
        cmdSearch.SetFocus
    End If
End Sub

Private Sub cmdAdd_Click()
    pnlYYMM.Caption = cboYear.Text & "년 " & cboMonth.Text & "월"
    pnlCustom = Trim(grdOut.TextMatrix(grdOut.Row, 3))
    fraSearch.Enabled = False
    pnlSumUp.Enabled = False
    
    Call ClearAddData
    fraAdd.Visible = True
    txtData(0).SetFocus
    
    fraAdd.Move 5850, 4200
End Sub

Private Sub cmdAddCancel_Click()
    fraSearch.Enabled = True
    pnlSumUp.Enabled = True
    fraAdd.Visible = False
End Sub

Private Sub cmdAddSave_Click()
    Dim i%
    
    If Not CheckAddData Then
        Exit Sub
    End If
    If MsgBox("가공료 데이터를 입력하시겠습니까?", vbYesNo + vbQuestion, "입력저장 전 확인") = vbYes Then
        m_CustomID = grdOut.TextMatrix(grdOut.Row, 22)
        If AddData() Then
            MsgBox "해당 가공료 데이터를 입력하였습니다.", vbInformation + vbOKOnly, "입력처리 완료"
            fraSearch.Enabled = True
            pnlSumUp.Enabled = True
            fraAdd.Visible = False
            
            m_sOrderID = txtData(0).Tag
            m_nWorkUnitPrice = txtData(2)
            
            Call FillGridOut
            
            With grdOut
                For i = .FixedRows To .Rows - 1
                    If m_CustomID = .TextMatrix(i, 22) Then
                        .Row = i
                        .TopRow = .Row
                        Exit For
                    End If
                Next i
            End With
            
            With grdSumUp
                For i = .FixedRows To .Rows - .FixedRows
                    If m_sOrderID = .TextMatrix(i, 21) And m_nWorkUnitPrice = .TextMatrix(i, 29) Then
                        .Row = i
                        .TopRow = .Row
                    End If
                Next i
            End With
            
            Call GetMaxTaxSeq
            
        End If
    End If
    
End Sub

Private Sub cmdCancel_Click()
    fraDetail.Visible = False
End Sub

Private Sub ClearData()
Dim i%

    fraDetail.Visible = False

    pnlTerm.Caption = ""
    pnlTerm.Visible = False

    txtOrderNO = ""
    txtArticle = ""
    txtWorkName = ""
    txtOrderQty = ""
    txtPrevOutQty = ""
    txtPrevSumQty = ""
    
    pnlCount.Caption = ""
    pnlQty.Caption = ""
    pnlSumPrice.Caption = ""
    
    For i = 0 To 2
        txtAddTaxSeq(i) = ""
        txtAddTaxSeq(i).Tag = ""
    Next i
    pnlSumSupplyPrice(0).Caption = ""
    pnlAddTaxPrice.Caption = ""
    pnlAddTaxSum.Caption = ""
    
    txtUnitPrice = ""
    txtOutQty = ""
    txtSumQty = ""
    txtUnitClss = ""
    txtSupplyPrice = ""
    txtExchangeRate = ""
    txtForeignPrice = ""
    txtAmount = ""
    txtTax = ""
    chkFreeTax.Value = 0
    
'    cboDealClss.ListIndex = -1
'    cboAdjustClss.ListIndex = -1
'    cboCurrency.ListIndex = -1
    
    cmdCloseAndComp.Enabled = False
    cmdComplete.Enabled = False
    cmdUpdate.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    
    fraAdd.Visible = False
End Sub

Private Function CheckAddData() As Boolean
    CheckAddData = False
    If Trim(txtData(0).Tag) = "" Or Trim(txtData(0).Text) = "" Then
        MsgBox "Order No. 가 선택되지 않았거나 올바르지 않습니다", vbExclamation + vbOKOnly, "Order No 오류"
        Exit Function
    End If
    If Trim(txtData(1).Tag) = "" Or Trim(txtData(1).Text) = "" Then
        MsgBox "품명이 선택되지 않았거나 올바르지 않습니다", vbExclamation + vbOKOnly, "품명 오류"
        Exit Function
    End If
    If cboData(0).ListIndex < 0 Then
        MsgBox "가공구분이 선택되지 않았습니다", vbExclamation + vbOKOnly, "가공구분 오류"
        Exit Function
    End If
    If Not IsNumeric(txtData(2)) Then
        MsgBox "가공단가가 잘못되어 있습니다", vbExclamation + vbOKOnly, "가공단가 오류"
        Exit Function
    End If
    If cboData(5).ListIndex < 0 Then
        MsgBox "단위가 선택되지 않았습니다", vbExclamation + vbOKOnly, "단위선택 오류"
        Exit Function
    End If
    If cboData(1).ListIndex < 0 Then
        MsgBox "정산구분이 선택되지 않았습니다", vbExclamation + vbOKOnly, "정산구분 오류"
        Exit Function
    End If
    If Not IsNumeric(txtData(3)) Then
        MsgBox "계상수량이 잘못되어 있습니다", vbExclamation + vbOKOnly, "계상수량 오류"
        Exit Function
    End If
    If cboData(2).ListIndex < 0 Then
        MsgBox "거래구분이 선택되지 않았습니다", vbExclamation + vbOKOnly, "거래구분 오류"
        Exit Function
    End If
    If cboData(3).ListIndex < 0 Then
        MsgBox "부가세구분이 선택되지 않았습니다", vbExclamation + vbOKOnly, "부가세구분 오류"
        Exit Function
    End If
    If Not IsNumeric(txtData(4)) Then
        MsgBox "원화금액이 잘못되어 있습니다", vbExclamation + vbOKOnly, "원화금액 오류"
        Exit Function
    End If
    If cboData(4).ListIndex < 0 Then
        MsgBox "화폐단위가 선택되지 않았습니다", vbExclamation + vbOKOnly, "화폐단위 오류"
        Exit Function
    End If
    If cboData(4).ListIndex = 1 And Not (IsNumeric(txtData(5))) Then
        MsgBox "외화환율이 잘못되어 있습니다", vbExclamation + vbOKOnly, "외화환율 오류"
        Exit Function
    End If
    If cboData(4).ListIndex = 1 Then
        If Trim(txtData(6)) = "" Then
            MsgBox "외화금액이 잘못되어 있습니다", vbExclamation + vbOKOnly, "외화금액 오류"
            Exit Function
        ElseIf Not (IsNumeric(Mid(txtData(6), 2))) Then
            MsgBox "외화금액이 잘못되어 있습니다", vbExclamation + vbOKOnly, "외화금액 오류"
            Exit Function
        End If
    End If
    CheckAddData = True
End Function

Private Sub ClearAddData()
    Dim i%
    
    For i = 0 To txtData.Count - 1
        txtData(i).Text = ""
        txtData(i).Tag = ""
    Next i
    
    For i = 0 To cboData.Count - 1
        cboData(i).ListIndex = -1
    Next i
End Sub

Private Sub cmdCheck_Click(Index As Integer)
    Dim SetValue, i%
    
    
    If Index = 0 Or Index = 2 Then   '[0] 전체선택
        SetValue = flexChecked
    Else                '[1] 선택 해제
        SetValue = flexUnchecked
    End If

    If Index < 2 Then
        With grdOut
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 1) > 0 Then
                    .Cell(flexcpChecked, i, 1) = SetValue
                End If
            Next i
        End With
    Else
        With grdSumUp
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, 1) > 0 Then
                    .Cell(flexcpChecked, i, 1) = SetValue
                End If
            Next i
        End With
    End If
End Sub

Private Sub cmdCloseAndComp_Click()
    If MsgBox("현재 선택된 오더를 마감하시겠습니까?", vbYesNo + vbQuestion, "마감처리전 확인") = vbYes Then
        If CloseOrder() Then
            MsgBox "수주마감 처리를 하였습니다.", vbInformation + vbOKOnly, "수주 마감처리"
            Call FillGridOut
        End If
    End If
End Sub

Private Sub cmdComplete_Click()
    Dim i%
    If cmdComplete.Caption = "계상완료" Then
        If MsgBox("현재 선택된 거래처의 계상완료를 하시겠습니까?", vbYesNo + vbQuestion, "계상완료 확인") = vbYes Then
            m_CustomID = grdOut.TextMatrix(grdOut.Row, 22)
            If CompleteData(0) Then
                MsgBox "계상완료 처리를 하였습니다.", vbInformation + vbOKOnly, "계상 완료처리"
                Call FillGridOut
                
                With grdOut
                    For i = .FixedRows To .Rows - 1
                        If m_CustomID = .TextMatrix(i, 22) Then
                            .Row = i
                            .TopRow = .Row
                            Exit For
                        End If
                    Next i
                End With
            End If
        End If
    Else
        If MsgBox("현재 선택된 거래처의 계상완료를 하시않겠습니까?", vbYesNo + vbQuestion, "계상미완료 확인") = vbYes Then
            m_CustomID = grdOut.TextMatrix(grdOut.Row, 22)
            If CompleteData(1) Then
                MsgBox "계상미완료 처리를 하였습니다.", vbInformation + vbOKOnly, "계상 미완료처리"
                Call FillGridOut
                
                With grdOut
                    For i = .FixedRows To .Rows - 1
                        If m_CustomID = .TextMatrix(i, 22) Then
                            .Row = i
                            .TopRow = .Row
                            Exit For
                        End If
                    Next i
                End With
            End If
        End If
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim oSubul As PlusLib2.CSubul
    Dim TRec As PlusLib2.TProcCostDet
    Dim i%
    On Error GoTo ErrHandler
    
    If MsgBox("해당 청구건을 삭제하시겠습니까", vbYesNo) = vbYes Then
            Set oSubul = New PlusLib2.CSubul
            oSubul.Connection = g_adoCon
            oSubul.UserName = g_sUserName
            
        With TRec
            .sBasisYearMon = grdSumUp.TextMatrix(grdSumUp.Row, 27)
            .sCustomID = grdSumUp.TextMatrix(grdSumUp.Row, 28)
            .nProcSeq = grdSumUp.TextMatrix(grdSumUp.Row, 38)
        End With
        
        m_CustomID = grdOut.TextMatrix(grdOut.Row, 22)
        If oSubul.DeleteProcCostDet(TRec) Then
            Call FillGridOut
            
            With grdOut
                For i = .FixedRows To .Rows - 1
                    If m_CustomID = .TextMatrix(i, 22) Then
                        .Row = i
                        .TopRow = .Row
                        Exit For
                    End If
                Next i
            End With
            
        End If
        Set oSubul = Nothing
    End If
    
    Exit Sub
    
ErrHandler:
    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmProcCost.CmdDelete", Err.Description)
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Select Case Index
        Case 0:
            Call ReturnCode(LG_ORDER, , False, txtData(0))
            
            If Len(txtData(0).Tag) > 0 Then
                Call GetOrderOne(txtData(0).Tag)
            End If
            txtData(1).SetFocus
        Case 1:
            Call ReturnCode(LG_ARTICLE, , False, txtData(1))
            cboData(0).SetFocus
        Case 2:
            Call ReturnCode(LG_CUSTOM, 0, , txtSearch(1))
            cmdSearch.SetFocus
    End Select
End Sub

Private Sub cmdPrint_Click()
    Dim i%
    Dim sPrinter As String
        
    sPrinter = Printer.DeviceName
        
    If frmPrinter.SelectPrinter(sPrinter, m_sPrinter) Then
        Call PrintTax
        Call ReturnPrinter(sPrinter)
        
        m_sOrderID = grdSumUp.TextMatrix(grdSumUp.Row, 21)
        m_nWorkUnitPrice = txtUnitPrice
        
        Call FillGridSumUp
        
        With grdSumUp
            For i = .FixedRows To .Rows - 1
                If m_sOrderID = .TextMatrix(i, 21) And m_nWorkUnitPrice = .TextMatrix(i, 29) Then
                    .Row = i
                    .TopRow = .Row
                End If
            Next i
        End With
    End If
End Sub

Private Sub PrintTax()
    Dim oCustom As PlusLib2.CCustom
    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%, nFormulas%, nCnt%
    Dim nTQty&, nTAmount&, nTTax&
    Dim sOrderFlag$, sTaxClss$, sDealClss$
    
    On Error GoTo ErrHandler
    
    Set oCustom = New PlusLib2.CCustom
    oCustom.Connection = g_adoCon
    Set rs = oCustom.GetCustomOne(grdOut.TextMatrix(grdOut.Row, 22))
    Set oCustom = Nothing
    
    With cryReport
        .Reset
        .PrintFileType = crptText
        .ReportFileName = App.Path & "\Report\Tax.Rpt"
        '***************************************************************************
         '공급 받는자 정보 출력
        '---------------------------------------------------------------------------
        .Formulas(0) = "TaxSeq='" & Right(txtAddTaxSeq(0), 1) & txtAddTaxSeq(1) & "-" & txtAddTaxSeq(2) & "'"
        .Formulas(1) = "CustomNo= '" & Left(CheckNull(rs!CustomNo), 3) & " - " & Mid(CheckNull(rs!CustomNo), 4, 2) & " - " & Right(CheckNull(rs!CustomNo), 5) & "'"
        .Formulas(2) = "Custom='" & CheckNull(rs!kCustom) & "'"
        .Formulas(3) = "Chief='" & CheckNull(rs!Chief) & "'"
        ''        'S_201312_태을염직_99 에 의한 수정-OLD소스
''        .Formulas(4) = "Address='" & CheckNull(rs!Address1) & " " & CheckNull(rs!Address2) & "'"
        'S_201312_태을염직_99 에 의한 수정-NEW 소스
        If CheckNull(rs!Address1) <> "" Then             '도로명 주소 있으면
            .Formulas(4) = "Address='" & CheckNull(rs!Address1) & " " & CheckNull(rs!Address2) & "'"
        Else                            '도로명 주소 없으면-지번주소
            .Formulas(4) = "Address='" & CheckNull(rs!AddressJiBun1) & " " & CheckNull(rs!AddressJiBun2) & "'"
        End If
        .Formulas(5) = "Condition='" & CheckNull(rs!Condition) & "'"
        .Formulas(6) = "Category='" & CheckNull(rs!Category) & "'"
        '***************************************************************************
        
        'S_201312_태을염직_99 에 의한 추가-엑셀 하드 코딩 대신 DB에서 가져옴
        '***************************************************************************
        '공급자 정보 출력
        '---------------------------------------------------------------------------
        .ParameterFields(0) = "CustomNo1" & ";" & Format(g_companyInfo.Company_No, "###-##-#####") & ";True"                 '사업자번호
        .ParameterFields(1) = "Custom1" & ";" & g_companyInfo.Company_Name & ";True"                   '상호
        .ParameterFields(2) = "Chief1" & ";" & g_companyInfo.Chief & ";True"                    '대표자
        If CheckNull(g_companyInfo.Address1) <> "" Then              '도로명 주소 있으면
            .ParameterFields(3) = "Address1" & ";" & g_companyInfo.Address1 & " " & g_companyInfo.Address2 & ";True"                  '주소
        Else                            '도로명 주소 없으면-지번주소
            .ParameterFields(3) = "Address1" & ";" & g_companyInfo.AddressJiBun1 & " " & g_companyInfo.AddressJiBun2 & ";True"                  '주소
        End If

        .ParameterFields(4) = "Condition1" & ";" & g_companyInfo.Company_type & ";True"                '업태
        .ParameterFields(5) = "Category1" & ";" & g_companyInfo.Category & ";True"              '종목
        '***************************************************************************
        
    End With
    rs.Close
    Set rs = Nothing
        
        
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    Set rs = oSubul.GetTax(txtAddTaxSeq(0) & txtAddTaxSeq(1) & txtAddTaxSeq(2), MakeDate(DF_SHORT, dtpPrnDate))
    Set oSubul = Nothing
    
    cryReport.Formulas(7) = "PrnDate='" & rs!PrnDate & "'"
    
    sOrderFlag = rs!OrderFlag
    sTaxClss = rs!TaxClss
    sDealClss = rs!DealClss
    With cryReport
        For i = 0 To rs.RecordCount - 1
            nTQty = nTQty + rs!SumQty
            nTAmount = nTAmount + rs!Amount
            nTTax = nTTax + rs!Tax
            
            rs.MoveNext
        Next i
    End With
    rs.MoveFirst
    
    With cryReport
        nCnt = 0
        nFormulas = 7
        For i = 0 To rs.RecordCount - 1
            If nCnt < 3 Then
                .Formulas(nFormulas + (i * 5) + 1) = "Article" & (i + 1) & "='" & rs!Article & "'"
                .Formulas(nFormulas + (i * 5) + 2) = "WorkName" & (i + 1) & "='" & rs!WorkName & "'"
                .Formulas(nFormulas + (i * 5) + 3) = "SumQty" & (i + 1) & "='" & rs!SumQty & "'"
                .Formulas(nFormulas + (i * 5) + 4) = "Amount" & (i + 1) & "='" & rs!Amount & "'"
                .Formulas(nFormulas + (i * 5) + 5) = "Tax" & (i + 1) & "='" & IIf(rs!Tax = 0, "", rs!Tax) & "'"
            Else
                .Formulas(nFormulas + (i * 5) + 1) = "Article" & (i + 1) & "='외 " & rs.RecordCount - nCnt & "건'"
                Exit For
            End If
            nCnt = nCnt + 1
            rs.MoveNext
        Next i
    End With
    
    With cryReport
        .Formulas(28) = "TAmount='" & nTAmount & "'"
        .Formulas(29) = "TTax='" & nTTax & "'"
        .Formulas(30) = "Space='" & 10 - Len(CStr(nTAmount)) & "'"
        .Formulas(31) = "Total='" & nTAmount + nTTax & "'"
        
        If sOrderFlag = "0" And sTaxClss = "불포함" Then
            If sDealClss = "1" Then
                .Formulas(32) = "Remark='LC/OPEN'"
            ElseIf sDealClss = "2" Then
                .Formulas(32) = "Remark='구매승인서'"
            ElseIf sDealClss = "3" Then
                .Formulas(32) = "Remark='임가공계약서'"
            End If
        Else
            .Formulas(32) = "Remark=''"
        End If
        .SelectionFormula = ""
        .PrinterDriver = m_sPrinter
        .PrinterName = m_sPrinter
       .PrinterPort = "LPT1:"
        .WindowState = crptMaximized
'        If bPreview Then
'           .Destination = crptToWindow
'        Else
            .Destination = crptToPrinter
'        End If
            .CopiesToPrinter = 2
        .Action = 1
    End With
    Exit Sub
    
ErrHandler:
    Set oCustom = Nothing
    Set oSubul = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "frmProcCost.CmdPrint_Click", Err.Description)
End Sub

Private Sub cmdSearch_Click()
    Call ClearData
    Call FillGridOut
    Call GetMaxTaxSeq
End Sub


Private Sub cmdSave_Click()
    Dim i%
    
    If MsgBox("선택한 항목들에 대해서 계상처리를 하시겠습니까?", vbYesNo + vbQuestion, "최종 확인") = vbYes Then
        m_CustomID = grdOut.TextMatrix(grdOut.Row, 22)
        If SaveData() Then
            MsgBox "해당 건의 계상처리를 완료하였습니다", vbInformation + vbOKOnly, "계상처리 완료"
            Call FillGridOut
            
            With grdOut
                For i = .FixedRows To .Rows - 1
                    If m_CustomID = .TextMatrix(i, 22) Then
                        .Row = i
                        .TopRow = .Row
                        Exit For
                    End If
                Next i
            End With
            
            With grdSumUp
                For i = .FixedRows To .Rows - 1
                    If m_sOrderID = .TextMatrix(i, 21) And m_nWorkUnitPrice = .TextMatrix(i, 29) Then
                        .Row = i
                        .TopRow = .Row
                    End If
                Next i
            End With
        End If
    End If
    Call GetMaxTaxSeq
    fraDetail.Visible = False
End Sub

Private Function SaveData() As Boolean
    Dim oSubul As PlusLib2.CSubul
    Dim tWork() As PlusLib2.TProcCost
    Dim i%, iCntChk%, iCount%
    Dim sPrevMon$
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    SaveData = False

    sPrevMon = Format(DateAdd("m", -1, CDate(cboYear.Text & "-" & cboMonth.Text & "-" & "01")), "YYYYMM")
    With grdOut
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
                iCntChk = iCntChk + 1
            End If
        Next i
        
        If iCntChk = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "계상처리할 항목을 체크 선택한후 작업하여 주십시요", vbInformation + vbOKOnly, "항목 선택 요망"
            Exit Function
        End If
        
        Set oSubul = New PlusLib2.CSubul
        oSubul.Connection = g_adoCon
        oSubul.UserName = g_sUserName
        
        ReDim tWork(iCntChk)
        
        iCount = 0
        
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
            
                tWork(iCount).YearMon = cboYear.Text & cboMonth.Text
                tWork(iCount).CustomID = .TextMatrix(i, 22)
                tWork(iCount).AdjustClss = IIf(optAccount(0).Value = True, "오더", "출고")
                tWork(iCount).FromDate = Format(dtpDate(0), "YYYYMMDD")
                tWork(iCount).ToDate = Format(dtpDate(1), "YYYYMMDD")
                tWork(iCount).TaxClss = "포함"
                tWork(iCount).PrevYearMon = sPrevMon
                tWork(iCount).ExchRate = IIf(CheckNum(txtExchRate) = 0, 1, CheckNum(txtExchRate))
                iCount = iCount + 1
            End If
        Next i
    End With
    
    
    If Not oSubul.UpdateProcCost(tWork()) Then
        Set oSubul = Nothing
        SaveData = False
        Exit Function
    End If
    
    SaveData = True
    
    Set oSubul = Nothing

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    SaveData = False

    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmProcCost.SaveData", Err.Description)
End Function

Private Function UpdateData() As Boolean
    Dim oSubul As PlusLib2.CSubul
    Dim tItem As PlusLib2.TProcCostDet
    Dim rs As ADODB.Recordset
    Dim nForeign As Single
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    UpdateData = False

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    oSubul.UserName = g_sUserName
                    
    With grdSumUp
'        If CheckNum(txtUnitPrice) <> CheckNum(.TextMatrix(.Row, 29)) Then
'            Set rs = oSubul.GetProcCostDetailData(.TextMatrix(.Row, 27), .TextMatrix(.Row, 28), Trim(.TextMatrix(.Row, 21)), _
'                                    .TextMatrix(.Row, 22), .TextMatrix(.Row, 36), .TextMatrix(.Row, 23), CheckNum(txtUnitPrice))
'            If rs.RecordCount > 0 Then
'                MsgBox "변경되어질 단가는 이미 존재하는 단가입니다" & vbCrLf & _
'                        "다른 단가를 입력하여 주십시요", vbInformation + vbOKOnly, "단가 중복"
'                Set rs = Nothing
'                Set oSubul = Nothing
'                Exit Function
'            End If
'            Set rs = Nothing
'        End If
        
        
        If Trim(txtForeignPrice) <> "" Then
            nForeign = CSng(Mid(txtForeignPrice, 2))
        Else
            nForeign = 0
        End If
        
        
        If chkFreeTax.Value = vbChecked And cboDealClss.ItemData(cboDealClss.ListIndex) = 0 Then
            MsgBox ("영세율 선택시에는 결재구분을 반드시 선택하십시오. ")
            UpdateData = False
            Exit Function
        End If
        
        
        tItem.sBasisYearMon = .TextMatrix(.Row, 27)
        tItem.sCustomID = .TextMatrix(.Row, 28)
        tItem.nProcSeq = .TextMatrix(.Row, 38)
        tItem.sOrderNO = .TextMatrix(.Row, 7)
        tItem.sOrderID = .TextMatrix(.Row, 21)
        tItem.sArticleID = .TextMatrix(.Row, 22)
        tItem.sSubulWidthID = .TextMatrix(.Row, 36)
        tItem.sWorkID = .TextMatrix(.Row, 23)
        tItem.nWorkUnitPrice = CheckNum(txtUnitPrice)
        tItem.nTempUnitPrice = CheckNum(.TextMatrix(.Row, 29))
        tItem.nSumQty = CLng(txtOutQty)
        tItem.nSumQtyY = CLng(txtSumQty)
        tItem.sTaxClss = IIf(chkFreeTax.Value = 0, "포함", "불포함")
        tItem.sDealClss = Format(cboDealClss.ItemData(cboDealClss.ListIndex), "0")
        tItem.sAdjustClss = cboAdjustClss.Text
        tItem.sPriceClss = Format(cboCurrency.ItemData(cboCurrency.ListIndex), "0")
        tItem.nExchRate = CheckNum(txtExchangeRate)
        tItem.nAmountWon = CheckNum(txtAmount)
        tItem.nTax = CheckNum(txtTax)
        tItem.nTotalPrice = CheckNum(txtSupplyPrice)
        tItem.nAmountDollar = nForeign
        tItem.sOrderFlag = Format(CboOrderFlag.ItemData(CboOrderFlag.ListIndex), "0")
        
        If Not oSubul.UpdateProcCostDet(tItem) Then
            Set oSubul = Nothing
            UpdateData = False
            Exit Function
        End If
    End With
    
    UpdateData = True
    
    Set oSubul = Nothing

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    UpdateData = False

    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmProcCost.UpdateData", Err.Description)
End Function

Private Function AddData() As Boolean
    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim TRec As PlusLib2.TProcCostDet
    Dim nForeign As Single
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    AddData = False

    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    oSubul.UserName = g_sUserName

'    Set rs = oSubul.GetProcCostDetailData(grdOut.TextMatrix(grdOut.Row, 21), grdOut.TextMatrix(grdOut.Row, 22), _
'                        Trim(txtData(0).Tag), txtData(1).Tag, Format(cboData(7).ItemData(cboData(7).ListIndex), "00"), _
'                        Format(cboData(0).ItemData(cboData(0).ListIndex), "0000"), CSng(txtData(2).Text))
'    If rs.RecordCount > 0 Then
'        Set rs = Nothing
'        Set oSubul = Nothing
'        MsgBox "입력하고자 하는 가공료데이터는 이미 존재하는 건입니다." & vbCrLf & _
'                "다른 데이터로 변환하여 입력하여 주십시요", vbInformation + vbOKOnly, "데이터 중복"
'        Exit Function
'    End If
'    Set rs = Nothing
        
    If Trim(txtData(6)) <> "" Then
        nForeign = CSng(Mid(txtData(6), 2))
    Else
        nForeign = 0
    End If
    
    With TRec
        .sBasisYearMon = grdOut.TextMatrix(grdOut.Row, 21)
        .sCustomID = grdOut.TextMatrix(grdOut.Row, 22)
        .sOrderNO = Trim(txtData(0).Text)
        .sOrderID = txtData(0).Tag
        .sArticleID = txtData(1).Tag
        .sSubulWidthID = Format(cboData(7).ItemData(cboData(7).ListIndex), "00")
        .sWorkID = Format(cboData(0).ItemData(cboData(0).ListIndex), "0000")
        .nWorkUnitPrice = CSng(txtData(2).Text)
        .sAdjustClss = Trim(cboData(1).Text)
        .sUnitClss = Trim(cboData(5).Text)
        .nSumQty = CLng(txtData(3))
        .nSumQtyY = CLng(txtData(9))
        .nOrderQty = IIf(.sAdjustClss = "오더", .nSumQty, 0)
        .nOutQty = IIf(.sAdjustClss = "오더", 0, .nSumQty)
        .sTaxClss = Trim(cboData(3).Text)
        .nPrevMonSumQty = 0
        .sDealClss = Format(cboData(2).ItemData(cboData(2).ListIndex), "0")
        .sPriceClss = Format(cboData(4).ItemData(cboData(4).ListIndex), "0")
        .sOrderFlag = Format(cboData(6).ItemData(cboData(6).ListIndex), "0")
        If cboData(4).ListIndex = 1 Then
            .nExchRate = CSng(txtData(5))
        Else
            .nExchRate = 1
        End If
        .nAmountWon = CLng(txtData(4))
        .nAmountDollar = nForeign
        .nTax = txtData(8)
        .nTotalPrice = txtData(7)
    End With
    
    If Not oSubul.AddNewProcCostDet(TRec) Then
        Set oSubul = Nothing
        AddData = False
        Exit Function
    End If
    Set oSubul = Nothing
    AddData = True
    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    Set oSubul = Nothing
    AddData = False
    Call ErrorBox(Err.Number, "frmProcCost.AddData", Err.Description)
End Function

Private Function CompleteData(Index As Integer) As Boolean
    Dim oSubul As PlusLib2.CSubul
    Dim tWork() As PlusLib2.TMonthCustom
    Dim i%, iCntChk%, iCount%
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    CompleteData = False

    With grdOut
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
                iCntChk = iCntChk + 1
            End If
        Next i
        
        If iCntChk = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "계상완료할 거래처를 체크선택한후 작업하여 주십시요", vbInformation + vbOKOnly, "항목 선택 요망"
            Exit Function
        End If
        
        Set oSubul = New PlusLib2.CSubul
        oSubul.Connection = g_adoCon
        oSubul.UserName = g_sUserName
        
        ReDim tWork(iCntChk)
        
        iCount = 0
    
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
                tWork(iCount).sYearMon = cboYear.Text & cboMonth.Text
                tWork(iCount).sCustomID = .TextMatrix(i, 22)
                
                iCount = iCount + 1
            End If
        Next i
    
    
        If Not oSubul.UpdateComplete(tWork(), Index) Then
            Set oSubul = Nothing
            CompleteData = False
            Exit Function
        End If
    End With
    
    CompleteData = True
    
    Set oSubul = Nothing

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    CompleteData = False

    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmProcCost.CompleteData", Err.Description)
End Function

Private Function CloseOrder() As Boolean
    Dim oOrder As PlusLib2.COrder
    Dim i%, iCntChk%, iCount%
    Dim sOrderID() As String
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    CloseOrder = False

                    
    With grdSumUp
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
                iCntChk = iCntChk + 1
            End If
        Next i
        
        If iCntChk = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "마감처리할 오더를 체크선택한후 작업하여 주십시요", vbInformation + vbOKOnly, "항목 선택 요망"
            Exit Function
        End If
        
        ReDim sOrderID(iCntChk - 1)
        
        iCount = 0
    
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = flexChecked Then
                sOrderID(iCount) = .TextMatrix(i, 21)
                iCount = iCount + 1
            End If
        Next i
    
        Set oOrder = New PlusLib2.COrder
        oOrder.Connection = g_adoCon
        oOrder.UserName = g_sUserName
        
        If Not oOrder.UpdateOrderClose(sOrderID, 1) Then
            Set oOrder = Nothing
            CloseOrder = False
            Exit Function
        End If
        
    End With
    
    CloseOrder = True
    
    Set oOrder = Nothing

    Screen.MousePointer = vbDefault

    Exit Function

ErrHandler:
    Screen.MousePointer = vbDefault
    CloseOrder = False

    Set oOrder = Nothing
    Call ErrorBox(Err.Number, "frmProcCost.CloseOrder", Err.Description)
End Function

Private Sub cmdSumUp_Click()
Dim dDate As Date

'    If Not ((cboYear.Text & cboMonth.Text = ThisMon) Or (cboYear.Text & cboMonth.Text = LastMon)) Then
'        MsgBox "금월이나 전월에 한해서만 계상처리가 가능합니다", vbExclamation + vbOKOnly, "계상처리 불가"
'        Exit Sub
'    End If
    
    dDate = CDate(cboYear.Text & "-" & cboMonth.Text & "-" & "01")
    dtpDate(0).Value = dDate
    dDate = DateAdd("m", 1, dDate)
    dtpDate(1).Value = DateSerial(Year(dDate), Month(dDate), 1 - 1)
    txtExchRate.Text = ""
    
    fraDetail.Visible = True
End Sub

Private Sub cmdUpdate_Click()
    Dim i%
    
    If MsgBox("해당 오더(" & txtOrderNO & ")의 변경된 " & vbCrLf & vbCrLf & _
                "단가와 수량, 부가세구분을 적용시키겠습니까?", vbYesNo + vbQuestion, "수정전 확인") = vbYes Then
        m_CustomID = grdOut.TextMatrix(grdOut.Row, 22)
        If UpdateData() Then
            MsgBox "해당 건을 수정하였습니다.", vbInformation + vbOKOnly, "수정처리 완료"
            m_sOrderID = grdSumUp.TextMatrix(grdSumUp.Row, 21)
            m_nWorkUnitPrice = txtUnitPrice
            
            Call FillGridOut
            
            With grdOut
                For i = .FixedRows To .Rows - 1
                    If m_CustomID = .TextMatrix(i, 22) Then
                        .Row = i
                        .TopRow = .Row
                        Exit For
                    End If
                Next i
            End With
            
            With grdSumUp
                For i = .FixedRows To .Rows - 1
                    If m_sOrderID = .TextMatrix(i, 21) And m_nWorkUnitPrice = .TextMatrix(i, 29) Then
                        .Row = i
                        .TopRow = .Row
                    End If
                Next i
            End With
            
            Call GetMaxTaxSeq
        End If
        Screen.MousePointer = 1
        
    End If

End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Load()
    Dim i%
    m_bloading = True
    Me.Move 0, 0, 15300, 9660
    cmdSearch.Picture = LoadResPicture("SEARCH", vbResIcon)
    cmdSumUp.Picture = LoadResPicture("COMMAND", vbResIcon)

    ThisMon = Format(Now, "YYYYMM")
    LastMon = Format(DateAdd("m", -1, Now), "YYYYMM")
    
    For i = 0 To cmdFind.Count - 1
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
    Next i
    txtSearch(1).Enabled = False
    
    Call InitGrid
    Call SetOperate(Me)
    Call SetComboBox

    Call ClearData
    

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub FillGridOut()
    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    Set rs = oSubul.GetProcCost(cboYear.Text & cboMonth.Text, IIf(chkSearch(1).Value, 1, 0), txtSearch(1).Tag)
    Set oSubul = Nothing
        
    With grdOut
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 1 To rs.RecordCount
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            
            .TextMatrix(.Rows - 1, 0) = CStr(i)
'            If rs!CompClss = "*" Then
'                .Cell(flexcpChecked, .Rows - 1, 1) = flexNoCheckbox
'            Else
'                .Cell(flexcpChecked, .Rows - 1, 1) = flexUnchecked
'            End If
            .TextMatrix(.Rows - 1, 2) = IIf(rs!CompClss = "*", "■", "")
            .TextMatrix(.Rows - 1, 3) = rs!kCustom
            .TextMatrix(.Rows - 1, 4) = ""
            .TextMatrix(.Rows - 1, 5) = ""
            .TextMatrix(.Rows - 1, 6) = Format(rs!OutQty, "##,##0")
            .TextMatrix(.Rows - 1, 7) = Format(rs!AmountWon, "##,##0")
            .TextMatrix(.Rows - 1, 8) = Format(rs!AmountDollar, "##,##0.00")
            
            .TextMatrix(.Rows - 1, 19) = CheckNull(rs!FromDate)
            .TextMatrix(.Rows - 1, 20) = CheckNull(rs!ToDate)
            .TextMatrix(.Rows - 1, 21) = rs!YearMon
            .TextMatrix(.Rows - 1, 22) = rs!CustomID
            .TextMatrix(.Rows - 1, 23) = rs!CompClss
            
            rs.MoveNext
        Next i
        
        If rs.RecordCount = 0 Then
            grdSumUp.Rows = grdSumUp.FixedRows
        End If
        rs.Close
        
        Set rs = Nothing
        
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If
    End With
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmProcCost.FillGridOut", Err.Description)
End Sub

Private Sub InitGrid()
    Dim iCol%
    
    With grdOut
        .Redraw = flexRDNone
        
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusSolid
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .WordWrap = False
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
        
        .Cols = 24:     .Rows = 3
        .FixedCols = 1: .FixedRows = 3
        
        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .RowHeight(2) = 500

        For iCol = 0 To .Cols - 1
            .ColWidth(iCol) = 0
            .FixedAlignment(iCol) = flexAlignCenterCenter
        Next iCol
        
        .TextMatrix(2, 0) = "":                     .ColWidth(0) = 300:         .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(2, 1) = "":                     .ColWidth(1) = 250:         .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(2, 2) = "계상" & vbCrLf & "완료":                 .ColWidth(2) = 450:         .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(2, 3) = "거래처":               .ColWidth(3) = 2000:        .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(2, 4) = "수량"
        .TextMatrix(2, 5) = "단가"
        .TextMatrix(2, 6) = "출고량":             .ColWidth(6) = 1000:    .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(2, 7) = "청구량":        .ColWidth(7) = 1300:    .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(2, 8) = "합계":               .ColWidth(8) = 0:    .ColAlignment(8) = flexAlignRightCenter
        
        .TextMatrix(2, 19) = "FromDate"
        .TextMatrix(2, 20) = "ToDate"
        .TextMatrix(2, 21) = "기준년월"
        .TextMatrix(2, 22) = "거래처코드"
        .TextMatrix(2, 23) = "계상완료구분"
        
        .ColDataType(1) = flexDTBoolean
        .MergeCells = flexMergeFixedOnly
        .Redraw = flexRDDirect
    End With
    
    With grdSumUp
        .Redraw = flexRDNone
        
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .ScrollBars = flexScrollBarVertical
        .ScrollTrack = True
        .WordWrap = False
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
        
        .Cols = 39:     .Rows = 3
        .FixedCols = 1: .FixedRows = 3
        
        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .RowHeight(2) = 500

        For iCol = 0 To .Cols - 1
            .ColWidth(iCol) = 0
            .FixedAlignment(iCol) = flexAlignCenterCenter
        Next iCol
        
        .TextMatrix(2, 0) = "":             .ColWidth(0) = 300:         .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(2, 1) = "":             .ColWidth(1) = 250:         .ColAlignment(1) = flexAlignCenterCenter
        .TextMatrix(2, 2) = "수주" & vbCrLf & "마감":       .ColWidth(2) = 450:         .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(2, 3) = "영세" & vbCrLf & "율":        .ColWidth(3) = 450:         .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(2, 4) = "정산" & vbCrLf & "구분":        .ColWidth(4) = 450:         .ColAlignment(4) = flexAlignCenterCenter
        .TextMatrix(2, 5) = "거래" & vbCrLf & "구분":        .ColWidth(5) = 0:         .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(2, 6) = "관리번호":    .ColWidth(6) = 0:        .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(2, 7) = "Order No.":    .ColWidth(7) = 1300:        .ColAlignment(7) = flexAlignLeftCenter
        .TextMatrix(2, 8) = "품명":         .ColWidth(8) = 1800:        .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(2, 9) = "가공구분":     .ColWidth(9) = 900:        .ColAlignment(9) = flexAlignLeftCenter
        .TextMatrix(2, 10) = "단가":         .ColWidth(10) = 500:         .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(2, 11) = "전월" & vbCrLf & "(청구)":    .ColWidth(11) = 1000:    .ColAlignment(11) = flexAlignRightCenter
        .TextMatrix(2, 12) = "금월" & vbCrLf & "(청구)":   .ColWidth(12) = 900:         .ColAlignment(12) = flexAlignRightCenter
        .TextMatrix(2, 13) = "금월" & vbCrLf & "(청구)":         .ColWidth(13) = 0:         .ColAlignment(13) = flexAlignLeftCenter
        .TextMatrix(2, 14) = "화폐" & vbCrLf & "단위":         .ColWidth(14) = 0:         .ColAlignment(14) = flexAlignCenterCenter
        .TextMatrix(2, 15) = "환율":         .ColWidth(15) = 0:         .ColAlignment(15) = flexAlignRightCenter
        .TextMatrix(2, 16) = "원화금액"
        .TextMatrix(2, 17) = "외화금액"
        .TextMatrix(2, 18) = "부가세"
        .TextMatrix(2, 19) = "청구금액"
        .TextMatrix(2, 20) = "계상수량Y"
        
        
        .TextMatrix(2, 21) = "OrderID"
        .TextMatrix(2, 22) = "ArticleID"
        .TextMatrix(2, 23) = "WorkID"
        .TextMatrix(2, 24) = "OrderQty"
        .TextMatrix(2, 25) = "OutQty"
        .TextMatrix(2, 26) = "PrevMonOutQty"
        .TextMatrix(2, 27) = "YearMon"
        .TextMatrix(2, 28) = "CustomID"
        .TextMatrix(2, 29) = "WorkUnitPrice"
        .TextMatrix(2, 30) = "DealClss"
        .TextMatrix(2, 31) = "PriceClss"
        .TextMatrix(2, 32) = "ExchRate"
        .TextMatrix(2, 33) = "TaxSeq"
        .TextMatrix(2, 34) = "PrnDate"
        .TextMatrix(2, 35) = "OrderFlag"
        .TextMatrix(2, 36) = "SubulWidthID"
        .TextMatrix(2, 37) = "계산서":  .ColWidth(37) = 1000:   .ColAlignment(37) = flexAlignCenterCenter
        .TextMatrix(2, 38) = "ProcSeq"
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(2) = True
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub SetComboBox()
Dim iCount As Integer
Dim dDate As Date
    
    dDate = DateAdd("m", -1, Now)
    dtpDate(0).Value = dDate
    dtpDate(1).Value = dDate
    
    With cboYear
        .Clear
        For iCount = 1 To 3
            .AddItem Year(Now) - iCount
        Next iCount
        .AddItem Year(Now)
        .Text = Year(dDate)
    End With

    With cboMonth
        .Clear
        For iCount = 1 To 12
            .AddItem Format(iCount, "00")
        Next iCount
        .Text = Format(Month(dDate), "00")
    End With
    
    ' 영세율의 구분
    With cboDealClss
        .Clear
        .AddItem "":                       .ItemData(0) = 0
        .AddItem "1. LC/OPEN":             .ItemData(1) = 1
        .AddItem "2. 구매승인서":          .ItemData(2) = 2
        .AddItem "3. 임가공계약서":        .ItemData(3) = 3
        .ListIndex = -1
    End With
    With cboData(2)
        .Clear
        .AddItem "1. 내수":          .ItemData(0) = 1
        .AddItem "2. Local":         .ItemData(1) = 3
        .AddItem "3. Driect":        .ItemData(2) = 5
        .ListIndex = -1
    End With
    
    ' 화폐구분
    With cboCurrency
        .Clear
        .AddItem "\":        .ItemData(0) = 0
        .AddItem "$":        .ItemData(1) = 1
    End With
    With cboData(4)
        .Clear
        .AddItem "\":        .ItemData(0) = 0
        .AddItem "$":        .ItemData(1) = 1
    End With
    
    With cboData(3)
        .Clear
        .AddItem "포함"
        .AddItem "불포함"
    End With
    
    With cboData(5)
        .Clear
        .AddItem "YD"
        .AddItem "MT"
    End With
    
    ' 정산료 구분
    With cboData(1)
        .Clear
        .AddItem "오더"
        .AddItem "출고"
    End With
    With cboAdjustClss
        .Clear
        .AddItem "오더"
        .AddItem "출고"
    End With
    
    With CboOrderFlag
        .AddItem "Local":         .ItemData(0) = 0
        .AddItem "내수":          .ItemData(1) = 1
        .AddItem "시가공":        .ItemData(2) = 2
        .AddItem "샘플":          .ItemData(3) = 3
    End With
    
    With cboData(6)
        .AddItem "0.비사용":         .ItemData(0) = 0
        .AddItem "1.사용":           .ItemData(1) = 1
        .ListIndex = 0
    End With
    
    With cboName(22)
        .AddItem "9. 전체"
        .AddItem "0. 비사용"
        .AddItem "1. 사용"
        .ListIndex = 0
    End With
    
    Call MakeCodeCombo(cboData(0), CD_WORK)        ' 가공 구분
    Call MakeCodeCombo(cboData(7), CD_WIDTH, , False)
End Sub

Private Sub grdOut_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdOut
        If Col = 1 Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub grdOut_RowColChange()
    If m_bloading = False Then Exit Sub
    
    Call FillGridSumUp
End Sub

Private Sub FillGridSumUp()
    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%
    Dim sYearMon$, sCustomID$
    Dim nCount As Integer
    Dim nQty As Long
    Dim nPrice As Long
    Dim nAddTaxPrice As Long
    Dim nFreeTaxPrice As Long
    Dim nTotalTax As Long
    
    On Error GoTo ErrHandler
    m_bLoading1 = False
    
    With grdOut
        If .Rows > .FixedRows And .Row >= .FixedRows Then
            sYearMon = Trim(.TextMatrix(.Row, 21))
            sCustomID = Trim(.TextMatrix(.Row, 22))
        End If
    End With
    
    With grdSumUp
        Call ClearData
        If sYearMon <> "" Then
            cmdCloseAndComp.Enabled = True
            cmdComplete.Enabled = True
            cmdUpdate.Enabled = True
            cmdAdd.Enabled = True
            cmdDelete.Enabled = True

            Set oSubul = New PlusLib2.CSubul
            oSubul.Connection = g_adoCon
            
            Set rs = oSubul.GetProcCostDetail(sYearMon, sCustomID, Left(cboName(22).Text, 1))
            Set oSubul = Nothing
        
            .Redraw = flexRDBuffered
            .Rows = .FixedRows
            
            nCount = rs.RecordCount
            If rs.RecordCount > 0 Then
                pnlTerm = "기간: " & Format(rs!FromDate, "0000/00/00") & " ~ " & Format(rs!ToDate, "0000/00/00")
                pnlTerm.Visible = True
                
                    txtAddTaxSeq(0) = Left(rs!AddTaxSeq, 2)
                    txtAddTaxSeq(0).Tag = rs!TaxSeq
                    txtAddTaxSeq(1) = Mid(rs!AddTaxSeq, 3, 2)
                    txtAddTaxSeq(2) = Right(rs!AddTaxSeq, 4)
                    
                    If Len(Trim(rs!PrnDate)) = 0 Then
                        dtpPrnDate = MakeDate(DF_FULL, Now)
                    Else
                        dtpPrnDate = MakeDate(DF_FULL, rs!PrnDate)
                    End If
                
                For i = 1 To rs.RecordCount
                    .Rows = .Rows + 1
                    .RowHeight(.Rows - 1) = 300
                    .TextMatrix(.Rows - 1, 0) = CStr(i)
                    If Trim(rs!CloseClss) = "" Then
                        .Cell(flexcpChecked, .Rows - 1, 1) = flexUnchecked
                    Else
                        .Cell(flexcpChecked, .Rows - 1, 1) = flexNoCheckbox
                    End If
                    .TextMatrix(.Rows - 1, 2) = IIf(Trim(rs!CloseClss) = "", "", "■")
                    .TextMatrix(.Rows - 1, 3) = IIf(rs!TaxClss = "포함", "", "■")
                    .TextMatrix(.Rows - 1, 4) = rs!AdjustClss
                    
                    Select Case rs!DealClss
                        Case "1":
                            .TextMatrix(.Rows - 1, 5) = "내수"
                        Case "3":
                            .TextMatrix(.Rows - 1, 5) = "Local"
                        Case "5":
                            .TextMatrix(.Rows - 1, 5) = "Direct"
                    End Select
                    .TextMatrix(.Rows - 1, 6) = MakeOrderID(rs!OrderID, OM_EXPAND)
                    .TextMatrix(.Rows - 1, 7) = Trim(rs!OrderNo)
                    .TextMatrix(.Rows - 1, 8) = MakeArticle(rs!Article, rs!SubulWidth)
                    .TextMatrix(.Rows - 1, 9) = Trim(rs!WorkName)
                    .TextMatrix(.Rows - 1, 10) = CheckNum(rs!WorkUnitPrice) 'Format(rs!WorkUnitPrice, "##,###.##")
                    
                    .TextMatrix(.Rows - 1, 11) = Format(rs!PrevMonSumQty, "#,##0")
                    .TextMatrix(.Rows - 1, 12) = Format(rs!SumQty, "##,##0")
                    .TextMatrix(.Rows - 1, 13) = IIf(rs!UnitClss = "YD", "Y", IIf(rs!UnitClss = "MT", "M", "K"))
                    Select Case rs!Priceclss
                        Case "0":
                            .TextMatrix(.Rows - 1, 14) = ""
                        Case "1":
                            .TextMatrix(.Rows - 1, 14) = "$"
                    End Select
                    .TextMatrix(.Rows - 1, 15) = IIf(rs!Priceclss = "0", "", Format(rs!ExchRate, "##,###.00"))
                    .TextMatrix(.Rows - 1, 16) = Format(rs!AmountWon, "#,###")
                    .TextMatrix(.Rows - 1, 17) = Format(rs!AmountDollar, "#,###.00")
                    .TextMatrix(.Rows - 1, 18) = Format(rs!Tax, "#,##0")
                    .TextMatrix(.Rows - 1, 19) = Format(rs!TotalPrice, "#,##0")
                    .TextMatrix(.Rows - 1, 20) = Format(rs!SumQtyY, "#,##0")

                    .TextMatrix(.Rows - 1, 21) = rs!OrderID
                    .TextMatrix(.Rows - 1, 22) = rs!ArticleID
                    .TextMatrix(.Rows - 1, 23) = rs!WorkID
                    .TextMatrix(.Rows - 1, 24) = Format(rs!OrderQty, "##,##0")
                    .TextMatrix(.Rows - 1, 25) = Format(rs!OutQty, "##,##0")
                    .TextMatrix(.Rows - 1, 26) = Format(rs!PrevMonOutQty, "##,##0")
                    .TextMatrix(.Rows - 1, 27) = rs!BasisYearMon
                    .TextMatrix(.Rows - 1, 28) = rs!CustomID
                    .TextMatrix(.Rows - 1, 29) = rs!WorkUnitPrice
                    .TextMatrix(.Rows - 1, 30) = rs!DealClss
                    .TextMatrix(.Rows - 1, 31) = rs!Priceclss
                    .TextMatrix(.Rows - 1, 32) = rs!ExchRate
                    .TextMatrix(.Rows - 1, 33) = rs!TaxSeq
                    .TextMatrix(.Rows - 1, 34) = rs!PrnDate
                    .TextMatrix(.Rows - 1, 35) = rs!OrderFlag
                    .TextMatrix(.Rows - 1, 36) = rs!SubulWidthID
                    .TextMatrix(.Rows - 1, 37) = MakeTaxSeq(rs!TaxSeq, OM_COMPACT)
                    .TextMatrix(.Rows - 1, 38) = rs!ProcSeq
                    
                    ' 나중에 Kg단위도 환산처리해야 함.
                    nQty = nQty + rs!SumQtyY
                    nPrice = nPrice + rs!TotalPrice
                    nTotalTax = nTotalTax + rs!Tax
                    
                    If rs!TaxClss = "포함" Then
                        nAddTaxPrice = nAddTaxPrice + rs!AmountWon
                    Else
                        nFreeTaxPrice = nFreeTaxPrice + rs!AmountWon
                    End If
                    rs.MoveNext
                Next i
                
            End If
            rs.Close
            Set rs = Nothing
            
            .Redraw = flexRDDirect
            If .Rows > .FixedRows Then
                .Row = .FixedRows
            End If
        Else
            .Rows = .FixedRows
        End If
        
        pnlCount = Format(nCount, "##0 ")
        pnlQty = Format(nQty, "##,##0 ")
        pnlSumPrice = Format(nPrice, "##,##0 ")
    End With
    
    m_bLoading1 = False
    
    Call ChangeMode
    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oSubul = Nothing
    m_bLoading1 = False
    Call ErrorBox(Err.Number, "frmProcCost.grdOut_RowColChange", Err.Description)
End Sub

Private Sub grdSumUp_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdSumUp
        If Col = 1 Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With

End Sub

Private Sub grdSumUp_RowColChange()
    If m_bLoading1 Then Exit Sub
    Call ShowData
End Sub

Private Sub ShowData()
    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim i%
    Dim sYearMon$, sCustomID$
    
    On Error GoTo ErrHandler
    
    m_bLoading2 = True
    With grdSumUp
        If optOrder(0).Value Then
            txtOrderNO = .TextMatrix(.Row, 7)
            pnlName(0) = "Order No."
        Else
            txtOrderNO = .TextMatrix(.Row, 6)
            pnlName(0) = "관리 번호"
        End If
        txtArticle = .TextMatrix(.Row, 8)
        txtWorkName = .TextMatrix(.Row, 9)
        txtOrderQty = .TextMatrix(.Row, 24)
        txtPrevOutQty = .TextMatrix(.Row, 26)
        txtPrevSumQty = .TextMatrix(.Row, 11)
        txtUnitPrice = .TextMatrix(.Row, 10)
        txtOutQty = .TextMatrix(.Row, 12)
        txtSumQty = .TextMatrix(.Row, 20)
        txtUnitClss = .TextMatrix(.Row, 13)
        txtExchangeRate = .TextMatrix(.Row, 15)
        
        cboDealClss.ListIndex = FindComboBox(cboDealClss, CLng("0" & .TextMatrix(.Row, 30)))  '주문형태
        cboCurrency.ListIndex = FindComboBox(cboCurrency, CLng("0" & .TextMatrix(.Row, 31)))  '화폐구분
        cboAdjustClss.Text = Trim(.TextMatrix(.Row, 4))
        CboOrderFlag.ListIndex = FindComboBox(CboOrderFlag, CLng(.TextMatrix(.Row, 35))) '오더구분
        txtAmount = .TextMatrix(.Row, 16)
        txtTax = .TextMatrix(.Row, 18)
        txtSupplyPrice = .TextMatrix(.Row, 19)
        
        If .TextMatrix(.Row, 31) = "1" Then
            txtForeignPrice = Format(CheckNum(.TextMatrix(.Row, 17)), "$#,##0.00")
        Else
            txtForeignPrice = ""
        End If
        
        If Trim(.TextMatrix(.Row, 3)) = "" Then
            chkFreeTax.Value = 0
        Else
            chkFreeTax.Value = 1
        End If
        
        If Trim(.TextMatrix(.Row, 2)) = "" Then
            cmdCloseAndComp.Enabled = True
        Else
            cmdCloseAndComp.Enabled = False
        End If
        
        txtAddTaxSeq(0) = Left(.TextMatrix(.Row, 33), 2)
        txtAddTaxSeq(0).Tag = .TextMatrix(.Row, 33)
        txtAddTaxSeq(1) = Mid(.TextMatrix(.Row, 33), 3, 2)
        txtAddTaxSeq(2) = Right(.TextMatrix(.Row, 33), 4)
        chkOrderFlag.Value = IIf(.TextMatrix(.Row, 35) = "1", vbChecked, vbUnchecked)
        
        If Len(Trim(.TextMatrix(.Row, 34))) = 0 Then
            dtpPrnDate = MakeDate(DF_FULL, Now)
        Else
            dtpPrnDate = MakeDate(DF_FULL, .TextMatrix(.Row, 34))
        End If
    End With
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    Set rs = oSubul.GetTaxTotal(grdSumUp.TextMatrix(grdSumUp.Row, 33))
    Set oSubul = Nothing
    
    If rs.RecordCount = 0 Then
        pnlSumSupplyPrice(0) = "0"
        pnlAddTaxPrice = "0"
        pnlAddTaxSum = "0"
    Else
        pnlSumSupplyPrice(0) = Format(rs!Amount, "##,##0 ")
        pnlAddTaxPrice = Format(rs!Tax, "##,##0 ")
        pnlAddTaxSum = Format(Int(rs!Amount + rs!Tax), "##,##0 ")
    End If
    rs.Close
    Set rs = Nothing
    m_bLoading2 = False
    Exit Sub
    
ErrHandler:
    Set rs = Nothing
    Set oSubul = Nothing
    m_bLoading2 = False
    Call ErrorBox(Err.Number, "frmProcCost.ShowData", Err.Description)
End Sub



Private Sub optOrder_Click(Index As Integer)
    If optOrder(0).Value Then
        With grdSumUp
            .ColWidth(6) = 0
            .ColWidth(7) = 1200
            
            pnlName(0).Caption = "Order No."
            txtOrderNO = .TextMatrix(.Row, 7)
        End With
    Else
        With grdSumUp
            .ColWidth(6) = 1200
            .ColWidth(7) = 0
            
            pnlName(0).Caption = "관리번호"
            txtOrderNO = .TextMatrix(.Row, 6)
        End With
    End If
End Sub

Private Sub txtAddTaxSeq_GotFocus(Index As Integer)
    Call WholeSelect(txtAddTaxSeq(Index))
End Sub

Private Sub txtAddTaxSeq_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim i%
    Dim sTmpTaxSeq$
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtAddTaxSeq(2) = Format(txtAddTaxSeq(2), "0000")
        sTmpTaxSeq = txtAddTaxSeq(0) + txtAddTaxSeq(1) + txtAddTaxSeq(2)

        If txtAddTaxSeq(0).Tag <> sTmpTaxSeq Then
            If MsgBox("선택한 오더건에 대해서 세금계산서번호를 바꾸시겠습니까?", vbYesNo) = vbYes Then
                If ExistTaxSeq(sTmpTaxSeq) Then
                    If MsgBox("이미 존재하는 세금계산서번호입니다. 같은 번호로 부여하시겠습니까?", vbYesNo) = vbYes Then
                        Call UpdateTaxSeq
                    End If
                Else
                    Call UpdateTaxSeq
                End If
                m_sOrderID = grdSumUp.TextMatrix(grdSumUp.Row, 21)
                m_nWorkUnitPrice = txtUnitPrice
                
                Call FillGridSumUp
                
                With grdSumUp
                    For i = .FixedRows To .Rows - 1
                        If m_sOrderID = .TextMatrix(i, 21) And m_nWorkUnitPrice = .TextMatrix(i, 29) Then
                            .Row = i
                            .TopRow = .Row
                        End If
                    Next i
                End With
                
                Call GetMaxTaxSeq
                
            End If
        End If
        
        Call NextFocus
    End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call NextFocus
    End If
End Sub

Private Sub txtData_GotFocus(Index As Integer)
    Call WholeSelect(txtData(Index))
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
   
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Index = 0 Then               '관리번호/ 오더번호
            
            If ReturnCode(LG_ORDER, , False, txtData(0)) Then
            
                If Len(txtData(0)) > 0 Then
                    Call GetOrderOne(txtData(0).Tag)
                End If
                
                KeyAscii = 0
                Call NextFocus
            End If
            Call NextFocus
        ElseIf Index = 1 Then           '품명 코드
            Call ReturnCode(LG_ARTICLE, , False, txtData(1))
            Call NextFocus
        ElseIf Index = 2 Or Index = 3 Or Index = 5 Or Index = 9 Then
            txtData(2) = Format(txtData(2), "#,##0.00")
            txtData(3) = Format(txtData(3), "#,##0")
            txtData(5) = Format(txtData(5), "#,##0.00")
            
            If cboData(5).Text = "YD" Then
                txtData(9) = Format(txtData(3), "#,##0")
            Else
                txtData(9) = Format(CheckNum(txtData(3)) * 1.0936, "#,##0")
            End If
            
            If cboData(4).Text = "$" Then
                txtData(4) = Format(CheckNum(txtData(9)) * CheckNum(txtData(2)) * CheckNum(txtData(5)), "##,##0")
                If cboData(3).Text = "포함" Then
                    txtData(8) = Format(Fix(CheckNum(txtData(4)) * 0.1), "##,##0")
                Else
                    txtData(8) = "0"
                End If
                txtData(7) = Format(CheckNum(txtData(4)) + CheckNum(txtData(8)), "#,##0")
            
                If Trim(txtData(5).Text) <> "" And Trim(txtData(5).Text) <> "0" Then
                    txtData(6) = Format(CheckNum(txtData(4)) / CheckNum(txtData(5)), "$#,##0.00")
                Else
                    txtData(6) = ""
                End If
            Else
                txtData(5) = ""
                txtData(6) = ""
                
                txtData(4) = Format(CheckNum(txtData(9)) * CheckNum(txtData(2)) * CheckNum(txtData(6)), "##,##0")
                If cboData(3).Text = "포함" Then
                    txtData(8) = Format(Fix(CheckNum(txtData(4)) * 0.1), "##,##0")
                Else
                    txtData(8) = "0"
                End If
                txtData(7) = Format(CheckNum(txtData(4)) + CheckNum(txtData(8)), "#,##0")
            End If
            Call NextFocus
        ElseIf Index = 4 Or Index = 7 Or Index = 8 Then
            Call NextFocus
            
        End If
    End If
End Sub

Private Sub GetOrderOne(sOrderID As String)
    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    m_bLoading2 = True
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    Set rs = oSubul.GetOrderOne(sOrderID, grdOut.TextMatrix(grdOut.Row, 19), grdOut.TextMatrix(grdOut.Row, 20))
    
    cboData(0).ListIndex = FindComboBox(cboData(0), CLng(rs!WorkID))
    txtData(1) = rs!Article
    txtData(1).Tag = rs!ArticleID
    txtData(2) = rs!UnitPrice
    txtData(3) = Format(rs!OutQty, "#,##0")
    txtData(3) = rs!OutQty
    txtData(10) = rs!OutQtyY
    txtData(10).Tag = rs!OrderQty
    txtData(9) = Format(rs!OutQtyY, "#,##0")

    cboData(1).ListIndex = 1
    cboData(4).ListIndex = rs!Priceclss
    cboData(5).ListIndex = rs!UnitClss
    cboData(7).ListIndex = FindComboBox(cboData(7), CLng(rs!SubulWidthID))
    cboData(6).ListIndex = 0
    rs.Close
    Set rs = Nothing
    m_bLoading2 = False
    
    Exit Sub
ErrHandler:
    m_bLoading2 = False
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmProcCost.GetOrderOne", Err.Description)
End Sub


Private Sub txtExchangeRate_GotFocus()
    Call WholeSelect(txtExchangeRate)
End Sub

Private Sub txtExchangeRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(txtExchangeRate.Text) Then
            KeyAscii = 0
            
            txtExchangeRate = Format(txtExchangeRate, "##,##0.00")
            If cboCurrency.Text = "$" Then
                txtAmount = Format(CheckNum(txtSumQty) * CheckNum(txtUnitPrice) * CheckNum(txtExchangeRate), "##,##0")
                If chkFreeTax.Value = 0 Then
                    txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "##,##0")
                Else
                    txtTax = "0 "
                End If
                txtSupplyPrice = Format(Fix(CLng(txtAmount) + CLng(txtTax)), "##,##0")
            
                If Trim(txtExchangeRate.Text) <> "" And Trim(txtExchangeRate.Text) <> "0.00" Then
                    txtForeignPrice = Format(CheckNum(txtAmount) / CSng(txtExchangeRate), "$#,##0.00")
                Else
                    txtForeignPrice = ""
                End If
            Else
                txtExchangeRate = ""
                txtForeignPrice = ""
            End If
            
            Call NextFocus
        End If
    End If
End Sub

Private Sub txtExchRate_GotFocus()
    Call WholeSelect(txtExchRate)
End Sub

Private Sub txtExchRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(txtExchRate.Text) Then
            txtExchRate = Format(txtExchRate, "##,##0.00")
            KeyAscii = 0
            Call NextFocus
        Else
            MsgBox "숫자만 입력하여 주십시요", vbExclamation + vbOKOnly, "입력오류"
        End If
    End If
End Sub

Private Sub txtForeignPrice_GotFocus()
    Call WholeSelect(txtForeignPrice)
End Sub

Private Sub txtForeignPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        txtForeignPrice = Format(txtForeignPrice, "$#,##0.00")
        Call NextFocus
    End If
End Sub

Private Sub txtFreeTaxSeq_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call NextFocus
    End If
End Sub

Private Sub txtOutQty_GotFocus()
    Call WholeSelect(txtOutQty)
End Sub

Private Sub txtOutQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        txtOutQty = Format(txtOutQty, "##,##0")
        If txtUnitClss = "Y" Then
            txtSumQty = Format(txtOutQty, "##,##0")
        Else
            txtSumQty = Format(txtOutQty * 1.0936, "#,##0")
        End If
        
        If cboCurrency.Text = "$" Then
            txtAmount = Format(CheckNum(txtSumQty) * CheckNum(txtUnitPrice) * CheckNum(txtExchangeRate), "##,##0")
            If chkFreeTax.Value = 0 Then
                txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "##,##0")
            Else
                txtTax = "0 "
            End If
            txtSupplyPrice = Format(Fix(CLng(txtAmount) + CLng(txtTax)), "##,##0")
        
            If Trim(txtExchangeRate.Text) <> "" And Trim(txtExchangeRate.Text) <> "0" Then
                txtForeignPrice = Format(CheckNum(txtAmount) / CheckNum(txtExchangeRate), "$#,##0.00")
            Else
                txtForeignPrice = ""
            End If
        Else
            txtExchangeRate = ""
            txtForeignPrice = ""
            
            txtAmount = Format(CheckNum(txtSumQty) * CheckNum(txtUnitPrice), "##,##0")
            If chkFreeTax.Value = 0 Then
                txtTax = Format(Fix(CheckNum(txtSumQty) * CheckNum(txtUnitPrice) * 0.1), "##,##0")
            Else
                txtTax = "0 "
            End If
            txtSupplyPrice = Format(Fix(CLng(txtAmount) + CLng(txtTax)), "##,##0")
        End If
        
        Call NextFocus
    End If
End Sub

Private Sub txtSumQty_GotFocus()
    Call WholeSelect(txtSumQty)
End Sub

Private Sub txtSumQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        txtSumQty = Format(txtSumQty, "##,##0")
        
        If cboCurrency.Text = "$" Then
            txtAmount = Format(CheckNum(txtSumQty) * CheckNum(txtUnitPrice) * CheckNum(txtExchangeRate), "##,##0")
            If chkFreeTax.Value = 0 Then
                txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "##,##0")
            Else
                txtTax = "0 "
            End If
            txtSupplyPrice = Format(Fix(CLng(txtAmount) + CLng(txtTax)), "##,##0")
        
            If Trim(txtExchangeRate.Text) <> "" And Trim(txtExchangeRate.Text) <> "0" Then
                txtForeignPrice = Format(CheckNum(txtAmount) / CheckNum(txtExchangeRate), "$#,##0.00")
            Else
                txtForeignPrice = ""
            End If
        Else
            txtExchangeRate = ""
            txtForeignPrice = ""
            
            txtAmount = Format(CheckNum(txtSumQty) * CheckNum(txtUnitPrice), "##,##0")
            If chkFreeTax.Value = 0 Then
                txtTax = Format(Fix(CheckNum(txtSumQty) * CheckNum(txtUnitPrice) * 0.1), "##,##0")
            Else
                txtTax = "0 "
            End If
            txtSupplyPrice = Format(Fix(CLng(txtAmount) + CLng(txtTax)), "##,##0")
        End If
        
        Call NextFocus
    End If
End Sub


Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If KeyAscii = vbKeyReturn Then
            Call ReturnCode(LG_CUSTOM, 0, , txtSearch(1))
            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub txtSupplyPrice_GotFocus()
    Call WholeSelect(txtSupplyPrice)
End Sub

Private Sub txtSupplyPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call NextFocus
    End If
End Sub

Private Sub txtTax_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtSupplyPrice = Format(CheckNum(txtAmount) + CheckNum(txtTax), "#,##0")
        
        Call NextFocus
    End If
End Sub

Private Sub txtUnitPrice_GotFocus()
    Call WholeSelect(txtUnitPrice)
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        txtUnitPrice = Format(txtUnitPrice, "##,##0.00")
        
        If cboCurrency.Text = "$" Then
            txtAmount = Format(CheckNum(txtSumQty) * CheckNum(txtUnitPrice) * CheckNum(txtExchangeRate), "##,##0")
            If chkFreeTax.Value = 0 Then
                txtTax = Format(Fix(CheckNum(txtAmount) * 0.1), "##,##0")
            Else
                txtTax = "0"
            End If
            txtSupplyPrice = Format(Fix(CheckNum(txtAmount) + CheckNum(txtTax)), "##,##0")
        
            If Trim(txtExchangeRate.Text) <> "" And Trim(txtExchangeRate.Text) <> "0" Then
                txtForeignPrice = Format(CheckNum(txtAmount) / CSng(txtExchangeRate), "$#,##0.00")
            Else
                txtForeignPrice = ""
            End If
        Else
            txtExchangeRate = ""
            txtForeignPrice = ""
            
            txtAmount = Format(CheckNum(txtSumQty) * CheckNum(txtUnitPrice), "##,##0")
            If chkFreeTax.Value = 0 Then
                txtTax = Format(Fix(CheckNum(txtSumQty) * CheckNum(txtUnitPrice) * 0.1), "##,##0")
            Else
                txtTax = "0"
            End If
            txtSupplyPrice = Format(Fix(CheckNum(txtAmount) + CheckNum(txtTax)), "##,##0")
            
        End If
        
        Call NextFocus
    End If
End Sub

Private Sub GetMaxTaxSeq()
    Dim oSubul As PlusLib2.CSubul
    Dim sTaxSeq$
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    sTaxSeq = oSubul.GetTaxSeqMax(cboYear.Text & cboMonth.Text)
    Set oSubul = Nothing
    
    pnlMaxTaxSeq = Left(sTaxSeq, 2) & "-" & Mid(sTaxSeq, 3, 2) & "-" & Right(sTaxSeq, 4)
End Sub

Private Function ExistTaxSeq(sTaxSeq As String) As Boolean
    Dim oSubul As PlusLib2.CSubul
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    If oSubul.GetTaxSeq(cboYear.Text & cboMonth.Text, sTaxSeq) Then
        ExistTaxSeq = True
    Else
        ExistTaxSeq = False
    End If
    Set oSubul = Nothing
End Function

Private Sub UpdateTaxSeq()
    Dim oSubul As PlusLib2.CSubul
    Dim tItem As PlusLib2.TProcCostDet
    
    With grdSumUp
        tItem.sBasisYearMon = .TextMatrix(.Row, 27)
        tItem.sCustomID = .TextMatrix(.Row, 28)
        tItem.nProcSeq = .TextMatrix(.Row, 38)
        tItem.sTaxSeq = txtAddTaxSeq(0) & txtAddTaxSeq(1) & txtAddTaxSeq(2)
    End With
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    oSubul.UserName = g_sUserName
    
    If oSubul.ModifyTaxSeq(tItem) Then
        Set oSubul = Nothing
        Exit Sub
    End If

    Set oSubul = Nothing
End Sub


Private Sub ChangeMode()
    If grdOut.TextMatrix(grdOut.Row, 2) = "■" Then
        pnlProcCost.Enabled = False
        cmdComplete.Caption = "계상미완료"
    Else
        pnlProcCost.Enabled = True
        cmdComplete.Caption = "계상완료"
    End If
End Sub

