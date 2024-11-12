VERSION 5.00
Begin VB.Form frmChange 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "암호 변경"
   ClientHeight    =   1620
   ClientLeft      =   5235
   ClientTop       =   2190
   ClientWidth     =   5730
   Icon            =   "frmChange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1620
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   2280
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취 소 (C)"
      Height          =   420
      Left            =   4155
      MouseIcon       =   "frmChange.frx":000C
      MousePointer    =   99  '사용자 정의
      TabIndex        =   5
      Top             =   570
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확 인 (O)"
      Height          =   420
      Left            =   4155
      MouseIcon       =   "frmChange.frx":068E
      MousePointer    =   99  '사용자 정의
      TabIndex        =   4
      Top             =   75
      Width           =   1425
   End
   Begin VB.TextBox txtConfirm 
      Height          =   300
      IMEMode         =   3  '사용 못함
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1230
      Width           =   2280
   End
   Begin VB.TextBox txtNewPass 
      Height          =   300
      IMEMode         =   3  '사용 못함
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2280
   End
   Begin VB.TextBox txtOldPass 
      Height          =   300
      IMEMode         =   3  '사용 못함
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   450
      Width           =   2280
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "사용자 ID "
      Height          =   180
      Left            =   135
      TabIndex        =   9
      Top             =   150
      Width           =   825
   End
   Begin VB.Label lblPassWord 
      AutoSize        =   -1  'True
      Caption         =   "새 암호 확인(&F)"
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label lblPassWord 
      AutoSize        =   -1  'True
      Caption         =   "새 암호(&N)"
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   930
      Width           =   885
   End
   Begin VB.Label lblPassWord 
      AutoSize        =   -1  'True
      Caption         =   "이전 암호(&O)"
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   540
      Width           =   1065
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim oPerson As PlusLib2.CPerson
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler

    If txtOldPass <> g_sPassword Then
        MsgBox LoadResString(220), vbInformation
        txtOldPass.SetFocus
        Exit Sub
    End If
    If txtNewPass <> txtConfirm Then
        MsgBox LoadResString(219), vbInformation
        txtNewPass.SetFocus
        Exit Sub
    End If
    
    Set oPerson = New PlusLib2.CPerson
    oPerson.Connection = g_adoCon
    oPerson.UserName = g_sUserName
    If oPerson.ChangePassword(txtUserName, txtNewPass) Then
        MsgBox LoadResString(221), vbInformation
        Unload Me
    Else
        MsgBox LoadResString(222), vbExclamation
    End If
    Set oPerson = Nothing
    Exit Sub

ErrHandler:
    MsgBox "[" & Err.Number & "] : " & Err.Description, vbCritical
    
    Set oPerson = Nothing
    Set rs = Nothing
End Sub


Private Sub Form_Load()
    Me.Icon = LoadResPicture("KEY", vbResIcon)

    cmdOK.MouseIcon = LoadResPicture("POINTER", vbResCursor)
    cmdCancel.MouseIcon = LoadResPicture("POINTER", vbResCursor)

    txtUserName = g_sUserName
End Sub


Private Sub txtConfirm_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtNewPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtOldPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveFocus(KeyCode)
End Sub
