VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmContract 
   ClientHeight    =   2745
   ClientLeft      =   8235
   ClientTop       =   5370
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   4305
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   930
      Top             =   1110
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   330
      TabIndex        =   2
      Top             =   240
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1085
      _Version        =   196609
      ForeColor       =   16777215
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "수출 물품 임가공 계약서 "
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   660
      Left            =   2220
      TabIndex        =   0
      Top             =   1770
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1164
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   660
      Left            =   540
      TabIndex        =   1
      Top             =   1770
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   1164
      _Version        =   196609
      Caption         =   "      인쇄(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REPORTFILE = "\Report\ContractReport.rpt"

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdPrint_Click()
    Dim bPreview As Boolean
    
    Me.PopupMenu PlusMDI.mnuPopup
 '   Call Prn
    If PlusMDI.PrintPreview Then
        CrystalReport1.Destination = crptToWindow
    Else
        CrystalReport1.Destination = crptToPrinter
    End If
    
    CrystalReport1.ReportFileName = App.Path & REPORTFILE
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = True

End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = True
    

End Sub

Private Sub Form_Deactivate()
    PlusMDI.pnlMenu.Visible = True

End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 4425, 3150

    Call SetOperate(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True

End Sub



