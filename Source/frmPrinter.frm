VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmPrinter 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "프린터 선택"
   ClientHeight    =   2940
   ClientLeft      =   4815
   ClientTop       =   3840
   ClientWidth     =   5190
   Icon            =   "frmPrinter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VSFlex7LCtl.VSFlexGrid grdPrint 
      Height          =   2085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5145
      _cx             =   9075
      _cy             =   3678
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
   Begin Threed.SSCommand cmdSelect 
      Height          =   690
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   1635
      _ExtentX        =   2884
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
      Caption         =   "      선택(&S)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   3510
      TabIndex        =   2
      Top             =   2160
      Width           =   1635
      _ExtentX        =   2884
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
      Caption         =   "      취소(&C)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bSelect As Boolean
Private m_sPrinter As String
Private m_sSelectPrinter As String

Public Function SelectPrinter(sPrinter As String, Optional sSelectPrinter As String) As Boolean

    cmdSelect.Picture = LoadResPicture("SELECT", vbResIcon)
    cmdExit.Picture = LoadResPicture("CANCEL", vbResIcon)
    
    Call InitGrid
    Call LoadPrint
    m_sPrinter = sPrinter
    
    Me.Show vbModal
    SelectPrinter = m_bSelect
    sSelectPrinter = m_sSelectPrinter
End Function

Private Sub InitGrid()
    With grdPrint
        .Redraw = flexRDNone
        .Cols = 1
        Call SetVSFlexGrid(grdPrint)
        .FixedCols = 0
        
        .TextArray(0) = "프린터명": .ColAlignment(0) = flexAlignLeftCenter
        .FixedAlignment(0) = flexAlignCenterCenter
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub LoadPrint()
    Dim dPrinter As Printer
    
    For Each dPrinter In Printers
        grdPrint.AddItem dPrinter.DeviceName
    Next
    
End Sub

Private Sub GetPrinter()
    Dim dPrinter As Printer
    
    For Each dPrinter In Printers
        If dPrinter.DeviceName = grdPrint.TextMatrix(grdPrint.Row, 0) Then
            Set Printer = dPrinter
            m_sSelectPrinter = grdPrint.TextMatrix(grdPrint.Row, 0)
            m_bSelect = True
            Exit For
        End If
    Next
End Sub

Private Sub cmdExit_Click()
    m_bSelect = False
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    Call GetPrinter
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim i%
    With grdPrint
        For i = .FixedRows To .Rows - .FixedRows
            If m_sPrinter = .TextMatrix(i, 0) Then
                .Row = i
                .TopRow = .Row
                .SetFocus
            End If
        Next i
    End With
End Sub

Private Sub grdPrint_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call GetPrinter
        Me.Hide
    End If
End Sub
