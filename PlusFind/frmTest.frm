VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "코드찾기 테스트 화면"
   ClientHeight    =   4215
   ClientLeft      =   2040
   ClientTop       =   1530
   ClientWidth     =   7905
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   540
      Index           =   4
      Left            =   3045
      TabIndex        =   10
      Top             =   2640
      Width           =   540
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   540
      Index           =   3
      Left            =   2235
      TabIndex        =   9
      Top             =   2580
      Width           =   540
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   540
      Index           =   2
      Left            =   1485
      TabIndex        =   8
      Top             =   2565
      Width           =   540
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   540
      Index           =   1
      Left            =   780
      TabIndex        =   7
      Top             =   2580
      Width           =   540
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   540
      Index           =   0
      Left            =   150
      TabIndex        =   6
      Top             =   2595
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   5595
      TabIndex        =   5
      Top             =   3000
      Width           =   2040
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2340
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   4128
      _Version        =   196609
      Caption         =   "SSPanel1"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel SSPanel2 
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   345
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         _Version        =   196609
         Caption         =   "코드 (Code)"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   4110
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2385
         TabIndex        =   1
         Top             =   360
         Width           =   4110
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   1050
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         _Version        =   196609
         Caption         =   "명칭 (Name)"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCon As ADODB.Connection

Private Enum ECodeFind
    cfCustom = 0
    cfArticle = 1
    cfperson = 2
    cfDefect = 3
    cfOrder = 4
    cfDyes = 5
    cfAux = 6
End Enum

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0: Call CodeFind(cfArticle)
        Case 1: Call CodeFind(cfAux)
        Case 2: Call CodeFind(cfCustom)
        Case 3: Call CodeFind(cfDefect)
        Case 4: Call CodeFind(cfOrder)
        Case 5: Call CodeFind(cfperson)
    End Select
End Sub

Private Sub Form_Load()
    Dim sConnect$
    
    sConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=WZServer;DATABASE=MRPPlus;"

    Set adoCon = New ADODB.Connection
    With adoCon
        .ConnectionString = sConnect
        .CursorLocation = adUseClient
        
        .Open sConnect
        
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    adoCon.Close
    
    Set adoCon = Nothing
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call CodeFind(cfCustom)
                
    End If
End Sub

Private Sub CodeFind(Index As ECodeFind)
    Dim bReturn As Boolean
    Dim oWizFind As WizFind.CCodeFind
    
    Set oWizFind = New WizFind.CCodeFind
    With oWizFind
        .Connection = adoCon
        
        bReturn = .Find(Index)
        If bReturn Then
            txtCode = .Data(0)
            txtName = .Data(1)
        End If
    End With
    Set oWizFind = Nothing
End Sub
