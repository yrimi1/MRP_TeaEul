VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmProcWorking 
   Caption         =   "������ �۾� ��Ȳ"
   ClientHeight    =   9435
   ClientLeft      =   -225
   ClientTop       =   825
   ClientWidth     =   15240
   Icon            =   "frmProcWorking.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   15240
   Begin VSFlex7LCtl.VSFlexGrid grdProcess 
      Height          =   8625
      Left            =   0
      TabIndex        =   4
      Top             =   660
      Width           =   15225
      _cx             =   26855
      _cy             =   15214
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   65280
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   0
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
      ExtendLastCol   =   -1  'True
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
   Begin Threed.SSFrame frmSearch 
      Height          =   645
      Left            =   8670
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   1138
      _Version        =   196609
      Begin Threed.SSPanel pnlOrder 
         Height          =   525
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   926
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.OptionButton optOrder 
            Caption         =   "Order No."
            Height          =   180
            Index           =   0
            Left            =   1590
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   180
            Width           =   1200
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "���� ��ȣ"
            Height          =   180
            Index           =   1
            Left            =   180
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   180
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   600
         Left            =   2940
         TabIndex        =   5
         Tag             =   "PERM_ADDNEW"
         Top             =   30
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   1058
         _Version        =   196609
         Caption         =   "        ��ȸ(&F)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   600
         Left            =   4740
         TabIndex        =   6
         Top             =   30
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   1058
         _Version        =   196609
         Caption         =   "        �ݱ�(&X)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame fraRefresh 
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1138
      _Version        =   196609
      Begin Threed.SSPanel pnlMsg 
         Height          =   525
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   926
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   12539970
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�׷��� ���� ����Դϱ�?"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmProcWorking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()
    Call FillGridOrder
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15360, 9840
    
    fraRefresh.Visible = False
    Call SetOperate(Me)
    Call InitGrid
    
    Call FillGridOrder
    
End Sub



Private Sub InitGrid()
    Dim i%
    
    With grdProcess
        .Redraw = flexRDNone
        
        .SelectionMode = flexSelectionFree
'        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .ScrollBars = flexScrollBarVertical
        
        .Rows = 4:          .Cols = 31
        .FixedRows = 4:     .FixedCols = 0
        
        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .RowHeight(2) = 0
        .RowHeight(3) = 400

        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 0
        Next i

        .TextMatrix(3, 0) = "":                 .ColWidth(0) = 0:       .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(3, 1) = "������":           .ColWidth(1) = 1100:    .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(3, 2) = "ȣ��":             .ColWidth(2) = 600:     .ColAlignment(2) = flexAlignCenterCenter
        .TextMatrix(3, 3) = "ī���ȣ":         .ColWidth(3) = 1400:    .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(3, 4) = "����":             .ColWidth(4) = 700:     .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(3, 5) = "����":             .ColWidth(5) = 800:     .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(3, 6) = "�ŷ�ó":           .ColWidth(6) = 1400:    .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(3, 7) = "ǰ��":             .ColWidth(7) = 2500:    .ColAlignment(7) = flexAlignLeftCenter
        .TextMatrix(3, 8) = "����":             .ColWidth(8) = 2000:    .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(3, 9) = "������ȣ":         .ColWidth(9) = 1500:    .ColAlignment(9) = flexAlignCenterCenter
        .TextMatrix(3, 10) = "Order No.":       .ColWidth(10) = 0:      .ColAlignment(10) = flexAlignLeftCenter
        .TextMatrix(3, 11) = "�۾���":          .ColWidth(11) = 800:    .ColAlignment(11) = flexAlignCenterCenter
        .TextMatrix(3, 12) = "������":          .ColWidth(12) = 0:      .ColAlignment(12) = flexAlignCenterCenter
        .TextMatrix(3, 13) = "���۽ð�":        .ColWidth(13) = 1200:    .ColAlignment(13) = flexAlignCenterCenter
        .TextMatrix(3, 14) = "�ҿ�ð�":        .ColWidth(14) = 1200:    .ColAlignment(14) = flexAlignCenterCenter
        .TextMatrix(3, 15) = "��������":        .ColWidth(15) = 0:      .ColAlignment(14) = flexAlignCenterCenter
        
        .TextMatrix(3, 20) = "�����ڵ��ڵ�"
        .TextMatrix(3, 21) = "ī���ȣ"
        .TextMatrix(3, 22) = "���ҹ�ȣ"
        .TextMatrix(3, 23) = "�ŷ�ó�ڵ�"
        .TextMatrix(3, 24) = "ǰ���ڵ�"
        .TextMatrix(3, 25) = "�����ڵ�"
        .TextMatrix(3, 26) = "������ȣ"
        .TextMatrix(3, 27) = "�۾����ڵ�"
        .TextMatrix(3, 28) = "������"
        .TextMatrix(3, 29) = "���۽ð�"
        .TextMatrix(3, 30) = "���������ڵ�"
        
        .Redraw = flexRDDirect
    End With
    

End Sub


Private Sub optOrder_Click(Index As Integer)
    With grdProcess
        If Index = 0 Then
            .ColWidth(9) = 0
            .ColWidth(10) = 1500
        Else
            .ColWidth(9) = 1500
            .ColWidth(10) = 0
        End If
    End With
End Sub

Private Sub cmdSearch_Click()
    Call FillGridOrder
End Sub

Private Sub FillGridOrder()
    Dim oCard As PlusLib2.CCard
    Dim rs As ADODB.Recordset
    Dim sYYMM$, sHHMM$
    Dim sProcMachine$
    Dim i%, iCardCnt%
    Dim dNow As Date
    
    On Error GoTo ErrHandler
    
    
    Set oCard = New PlusLib2.CCard
    oCard.Connection = g_adoCon
    
    dNow = Now
    
    sYYMM = Format(dNow, "YYYYMMDD")
    sHHMM = Format(dNow, "HHNN")
    
    pnlMsg.Caption = Format(dNow, "YYYY/MM/DD HH:NN") & " ���� ������ �۾����� ī�� ����Ʈ�Դϴ�"
    fraRefresh.Visible = True
    Set rs = oCard.GetCardWorking(sYYMM, sHHMM)
    Set oCard = Nothing
        
    With grdProcess
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 1 To rs.RecordCount
            If sProcMachine <> rs!ProcessID & rs!machineid Then
                iCardCnt = 1
        
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 300
                
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = rs!Process
                .TextMatrix(.Rows - 1, 2) = rs!machineid
                .TextMatrix(.Rows - 1, 3) = IIf(Len(Trim(rs!SplitID)) > 0, MakeCardID(rs!CardID, OM_EXPAND) & "(" & rs!SplitID & ")", MakeCardID(rs!CardID, OM_EXPAND))
                .TextMatrix(.Rows - 1, 4) = rs!workroll
                .TextMatrix(.Rows - 1, 5) = Format(rs!workqty, "##,##0")
                .TextMatrix(.Rows - 1, 6) = rs!kCustom
                .TextMatrix(.Rows - 1, 7) = rs!Article
                .TextMatrix(.Rows - 1, 8) = rs!Color
                .TextMatrix(.Rows - 1, 9) = MakeOrderID(rs!OrderID, OM_EXPAND)
                .TextMatrix(.Rows - 1, 10) = rs!OrderNo
                .TextMatrix(.Rows - 1, 11) = rs!Name
                .TextMatrix(.Rows - 1, 12) = rs!StartDate
                .TextMatrix(.Rows - 1, 13) = MakeDate(DF_MD, rs!StartDate) & " " & Format(rs!StartTime, "00:00")
                .TextMatrix(.Rows - 1, 14) = Format(rs!requiredtime, "##,##0") & " ��"
                .TextMatrix(.Rows - 1, 15) = "��������"
                
                .TextMatrix(.Rows - 1, 20) = rs!ProcessID
                .TextMatrix(.Rows - 1, 21) = rs!CardID
                .TextMatrix(.Rows - 1, 22) = rs!SplitID
                .TextMatrix(.Rows - 1, 23) = rs!CustomID
                .TextMatrix(.Rows - 1, 24) = rs!ArticleID
                .TextMatrix(.Rows - 1, 25) = rs!ColorID
                .TextMatrix(.Rows - 1, 26) = rs!OrderID
                .TextMatrix(.Rows - 1, 27) = rs!PersonID
                .TextMatrix(.Rows - 1, 28) = rs!StartDate
                .TextMatrix(.Rows - 1, 29) = rs!StartTime
                .TextMatrix(.Rows - 1, 30) = "���������ڵ�"
    
    
            Else
                iCardCnt = iCardCnt + 1
                .RowHeight(.Rows - 1) = 250 * iCardCnt
    
                .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 1, 3) & vbCrLf & IIf(Len(Trim(rs!SplitID)) > 0, MakeCardID(rs!CardID, OM_EXPAND) & "(" & rs!SplitID & ")", MakeCardID(rs!CardID, OM_EXPAND))
                .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 4) & vbCrLf & rs!workroll
                .TextMatrix(.Rows - 1, 5) = .TextMatrix(.Rows - 1, 5) & vbCrLf & Format(rs!workqty, "##,##0")
                .TextMatrix(.Rows - 1, 6) = .TextMatrix(.Rows - 1, 6) & vbCrLf & rs!kCustom
                .TextMatrix(.Rows - 1, 7) = .TextMatrix(.Rows - 1, 7) & vbCrLf & rs!Article
                .TextMatrix(.Rows - 1, 8) = .TextMatrix(.Rows - 1, 8) & vbCrLf & rs!Color
                .TextMatrix(.Rows - 1, 9) = .TextMatrix(.Rows - 1, 9) & vbCrLf & MakeOrderID(rs!OrderID, OM_EXPAND)
                .TextMatrix(.Rows - 1, 10) = .TextMatrix(.Rows - 1, 10) & vbCrLf & rs!OrderNo
            End If
            sProcMachine = rs!ProcessID & rs!machineid
            rs.MoveNext
        Next i
        rs.Close
        
        Set rs = Nothing
        
        .Redraw = flexRDDirect
'        .SetFocus
    End With
    
    
    
    Exit Sub

ErrHandler:
    Set oCard = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmProcWorking.FillGridOrder", Err.Description)
End Sub

