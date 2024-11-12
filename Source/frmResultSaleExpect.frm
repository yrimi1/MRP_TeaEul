VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmResultSaleExpect 
   ClientHeight    =   9255
   ClientLeft      =   240
   ClientTop       =   525
   ClientWidth     =   11850
   Icon            =   "frmResultSaleExpect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
   Begin Threed.SSPanel pnlSub 
      Height          =   2835
      Left            =   3120
      TabIndex        =   14
      Top             =   5010
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   5001
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel SSPanel4 
         Height          =   345
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   609
         _Version        =   196609
         ForeColor       =   16777215
         BackColor       =   16711680
         Caption         =   "검사성적서 내역"
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdOrder 
         Height          =   2415
         Left            =   30
         TabIndex        =   16
         Top             =   390
         Width           =   8085
         _cx             =   14261
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
      End
   End
   Begin Threed.SSPanel pnlProgress 
      Height          =   870
      Left            =   450
      TabIndex        =   6
      Top             =   3990
      Visible         =   0   'False
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   1535
      _Version        =   196609
      Alignment       =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
      Begin MSComctlLib.ProgressBar proProgress 
         Height          =   390
         Left            =   90
         TabIndex        =   7
         Top             =   375
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         Caption         =   "180"
         Height          =   180
         Left            =   195
         TabIndex        =   8
         Top             =   120
         Width           =   270
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   8370
      TabIndex        =   3
      Top             =   8430
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      발행(&P)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   10140
      TabIndex        =   2
      Top             =   8430
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   7545
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   11715
      _cx             =   20664
      _cy             =   13309
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
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
   Begin Threed.SSFrame frmSearch 
      Height          =   825
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   1455
      _Version        =   196609
      Begin VB.TextBox txtExchange 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7920
         TabIndex        =   10
         Top             =   315
         Visible         =   0   'False
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   435
         Left            =   2235
         TabIndex        =   5
         Top             =   210
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   767
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
         Format          =   116785152
         CurrentDate     =   37957
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   210
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   767
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "실적 일자"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   7950
         TabIndex        =   9
         Top             =   75
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "환율"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   690
         Left            =   10005
         TabIndex        =   11
         Top             =   60
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   1217
         _Version        =   196609
         Caption         =   "      검색(&F)"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdLeft 
         Height          =   435
         Left            =   1665
         TabIndex        =   12
         Top             =   210
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   767
         _Version        =   196609
         AutoSize        =   1
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdRight 
         Height          =   435
         Left            =   5205
         TabIndex        =   13
         Top             =   210
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   767
         _Version        =   196609
         AutoSize        =   1
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   60
      TabIndex        =   17
      Top             =   8430
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1138
      _Version        =   196609
      Caption         =   "상세내역 "
      Begin Threed.SSOption optView 
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   476
         _Version        =   196609
         Caption         =   "보임"
      End
      Begin Threed.SSOption optView 
         Height          =   270
         Index           =   1
         Left            =   1035
         TabIndex        =   19
         Top             =   255
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         _Version        =   196609
         Caption         =   "숨김"
         Value           =   -1
      End
   End
   Begin Threed.SSCommand cmdExcel 
      Height          =   690
      Left            =   6600
      TabIndex        =   20
      Top             =   8430
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1217
      _Version        =   196609
      Caption         =   "      엑셀(&E)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmResultSaleExpect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExcel_Click()
    Dim oExcel      As Excel.Application
    Dim oExcelBook  As Excel.Workbook
    Dim oExcelSheet As Excel.Worksheet
    Dim oFs         As FileSystemObject
    Dim i%, j%, sWeek$, sReport$
    Dim nBaseRow%, nMaxRow%, nRowCnt%, nCnt%
    
    On Error GoTo ErrHandler
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Set oExcel = New Excel.Application
    Set oExcelBook = oExcel.Workbooks.Open(App.Path & "\Report\ResultSaleExpect.xls")

    Set oFs = New FileSystemObject
    If Not oFs.FolderExists(App.Path & "\Excel") Then
        oFs.CreateFolder (App.Path & "\Excel")
    End If
    
    sReport = App.Path & "\Excel\예상매출_" & MakeDate(DF_SHORT, dtpDate) & ".xls"
    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
    Set oFs = Nothing

'        oExcel.WindowState = xlMaximized
'        oExcel.Application.Visible = True

    nMaxRow = 50
    With grdData
        
        oExcel.Cells(5, 2) = Format(dtpDate, "YYYY년 M월 D일")
        nBaseRow = 9

        For i = 5 To .Rows - 1
            If CheckNum(.TextMatrix(i, 0)) > 0 Then
                oExcel.Cells(nBaseRow + nRowCnt, 1) = .TextMatrix(i, 1) '거래처
                oExcel.Cells(nBaseRow + nRowCnt, 7) = .TextMatrix(i, 2) '품명
                oExcel.Cells(nBaseRow + nRowCnt, 16) = .TextMatrix(i, 4) '합격량
                oExcel.Cells(nBaseRow + nRowCnt, 26) = "@" & .TextMatrix(i, 6) '단가
                oExcel.Cells(nBaseRow + nRowCnt, 29) = .TextMatrix(i, 7) '금액
            Else
                oExcel.Cells(59 + nCnt, 16) = .TextMatrix(i, 4) '합격량
                oExcel.Cells(59 + nCnt, 29) = .TextMatrix(i, 7) '금액
                nCnt = nCnt + 1
            End If
            nRowCnt = nRowCnt + 1
        Next i
        
        If nMaxRow - nRowCnt <> 0 Then
            oExcel.Rows(nBaseRow + nRowCnt & ":" & nBaseRow + nRowCnt).Select
            oExcel.Rows(nBaseRow + nRowCnt & ":" & nBaseRow + nMaxRow - 1).Select
            oExcel.Selection.Delete Shift:=xlUp
        End If
    End With


    Call oExcelBook.SaveAs(sReport)

    If PlusMDI.PrintPreview Then
        oExcel.WindowState = xlMaximized
        oExcel.Application.Visible = True
    Else
        oExcel.ActiveWindow.SelectedSheets.PrintOut Copies:=1
        Call ProcessClose("XLMAIN")
    End If
    
    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    
    Exit Sub
    
ErrHandler:
    Call oExcelBook.SaveAs(sReport)
    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing

    Call ProcessClose("XLMAIN")
    Call ErrorBox(Err.Number, "frmResultSaleExpect.cmdExcel_Click", Err.Description)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLeft_Click()
    dtpDate = dtpDate - 1
    Call cmdSearch_Click
End Sub

Private Sub cmdPrint_Click()
    Dim i%
    Dim sDate As String, eDate As String, nPageHV As Integer
    
    
    With grdData
        .Redraw = flexRDNone
        .ExtendLastCol = False
        
        For i = 0 To 3
           .MergeRow(i) = True
        Next i
        

        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = "일 일 예 상 매 출 표"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpText, 3, 1, 3, 2) = "▶ 예상일자 : " & MakeDate(DF_FULL, dtpDate)
        
        Call SetPrintMode(grdData, 1, True, nPageHV)

        
''        For i = .Rows - 1 To .FixedRows Step -1
''            ' 일계, 총계의 금액은 BackColor을 설정 한다.
''            If .TextMatrix(i, .Cols - 1) = "3" Or .TextMatrix(i, .Cols - 1) = "2" Then
''
''                .Cell(flexcpBackColor, i, 2, i, 2) = PRNHeaderColor
''                .Cell(flexcpBackColor, i, 5, i, 5) = PRNHeaderColor
''                .Cell(flexcpAlignment, i, 1, i, 1) = flexAlignCenterCenter
''
''                .Cell(flexcpFontBold, i, 2, i, 2) = True
''                .Cell(flexcpFontBold, i, 5, i, 5) = True
''                .Cell(flexcpFontBold, i, 1, i, 1) = True
''
''                If .TextMatrix(i, .Cols - 1) = "2" Then
''                    Exit For
''                End If
''            End If
''
''        Next i
        
        .MergeCells = flexMergeFree
'        .ColHidden(0) = True
        .ColWidth(4) = 1250
        .ColWidth(6) = 900
        .ColWidth(7) = 1450
        
        .PrintGrid "태을염직", True, 1, 0, 500
        
'        .ColHidden(0) = False
        .ColWidth(4) = 1500
        .ColWidth(6) = 1200
        .ColWidth(7) = 1700
        
        Call SetPrintMode(grdData, 1, False)
        
        For i = .Rows - 1 To .FixedRows Step -1
            ' 일계, 총계의 금액은 BackColor을 설정 한다.
            If .TextMatrix(i, .Cols - 1) = "3" Or .TextMatrix(i, .Cols - 1) = "2" Then
                
'                .Cell(flexcpBackColor, i, 2, i, 2) = PRNHeaderColor
'                .Cell(flexcpBackColor, i, 5, i, 5) = PRNHeaderColor
                .Cell(flexcpAlignment, i, 1, i, 1) = flexAlignCenterCenter
                
                .Cell(flexcpFontBold, i, 2, i, 2) = True
                .Cell(flexcpFontBold, i, 5, i, 5) = True
                .Cell(flexcpFontBold, i, 1, i, 1) = True
                
                If .TextMatrix(i, .Cols - 1) = "2" Then
                    Exit For
                End If
            End If
            
        Next i
        
        .ExtendLastCol = True

        .Redraw = flexRDDirect
    End With


End Sub

Private Sub cmdRight_Click()
    dtpDate = dtpDate + 1
    Call cmdSearch_Click
End Sub

Private Sub cmdSearch_Click()
    Call FillGridData
End Sub

Private Sub dtpDate_CloseUp()
    Call cmdSearch_Click
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 11970, 9660
    
    Call SetOperate(Me)
    Call InitGrid
    dtpDate = Now
    
    cmdLeft.Picture = LoadResPicture("LEFT", vbResIcon)
    cmdRight.Picture = LoadResPicture("RIGHT", vbResIcon)
    
    pnlSub.Visible = False
    Call FillGridData
    pnlProgress.Visible = False
End Sub

Sub FillGrdOrder()
''    Dim Key_Var As Variant
''    Dim IOClss As String
''    Dim StuffDate As String, StuffClss As String, StuffSeq As Integer
''    Dim OrderID As String, OutSeq As Integer
''
''
''    Dim oSubul As Pluslib2.CSubul
''    Dim rs As Recordset, dCustom_str As String, dArticle_Str As String, dDate_str As String
''    Dim i%, sOrderID$, bFlag As Boolean, II%
''
''    Screen.MousePointer = vbHourglass
''
''   ' On Error GoTo ErrHandler
''
''    pnlSub.Visible = True
''    m_bLoading = True
''
''    Set oSubul = New Pluslib2.CSubul
''    oSubul.Connection = g_adoCon
''    IOClss = ""
''    StuffDate = "": StuffClss = "": StuffSeq = 0
''    OrderID = "": OutSeq = 0
''
''    With grdData(0)
''        Key_Var = Split(.TextMatrix(.Row, 17), "-")
''        IOClss = .TextMatrix(.Row, 16)
''    End With
''
''    If IOClss = "1" Then    ' 입고
''        StuffDate = Key_Var(0)
''        StuffClss = Key_Var(1)
''        StuffSeq = Key_Var(2)
''    Else
''        OrderID = Key_Var(0)
''        OutSeq = Key_Var(1)
''    End If
''
''    Set rs = oSubul.GetSubulOrderSub(IOClss, StuffDate, StuffClss, StuffSeq, OrderID, OutSeq)
''    Set oSubul = Nothing
''
''    With grdStuffIN
''        .Rows = .FixedRows
''    End With
''
''    With grdOutWare
''        .Rows = .FixedRows
''    End With
''
''    With grdOrder
''        .Rows = .FixedRows
''        Do Until rs.EOF
''            .AddItem "" & vbTab & rs!OrderNo & vbTab & SetCurrency(rs!OrderQty, 0) & vbTab & SetCurrency(rs!OutQty, 0) & vbTab & SetCurrency(rs!OrderQty - rs!OutQty) & vbTab & _
''                   SetCurrency(rs!StuffQty, 0) & vbTab & SetCurrency(rs!OutRealQty, 0) & vbTab & 0 & vbTab & rs!OrderID
''            rs.MoveNext
''        Loop
''        If .Rows > .FixedRows Then
''            .Row = .FixedRows
''            Call grdOrder_RowColChange
''        End If
''    End With
''    rs.Close
''    Set rs = Nothing
''
''    Screen.MousePointer = vbDefault
    
End Sub

Private Sub InitGrid()
    With grdData
        .Cols = 13
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .Rows = 5
        .FixedRows = 5
        
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .RowHidden(3) = True

        .TextMatrix(4, 0) = "":                   .ColWidth(0) = 400:         .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(4, 1) = "업체명":             .ColWidth(1) = 2300:        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(4, 2) = "ITEM":               .ColWidth(2) = 2500:        .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(4, 3) = "가공구분":           .ColWidth(3) = 1600:        .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(4, 4) = "합격량(Y)":          .ColWidth(4) = 1500:        .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(4, 5) = "불량량(Y)":          .ColWidth(5) = 0:           .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(4, 6) = "단가(@)":            .ColWidth(6) = 1200:        .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(4, 7) = "금액(\)":            .ColWidth(7) = 1700:        .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(4, 8) = "비고":               .ColWidth(8) = 0:           .ColAlignment(8) = flexAlignCenterCenter
        .TextMatrix(4, 9) = "CustomID":           .ColWidth(9) = 0:           .ColAlignment(9) = flexAlignCenterCenter
        .TextMatrix(4, 10) = "ArticleID":         .ColWidth(10) = 0:          .ColAlignment(10) = flexAlignCenterCenter
        .TextMatrix(4, 11) = "WorkID":            .ColWidth(11) = 0:          .ColAlignment(11) = flexAlignCenterCenter
        .TextMatrix(4, 12) = "sLevel":            .ColWidth(12) = 0:          .ColAlignment(12) = flexAlignCenterCenter
        
        
        .FontSize = 10
        .FontName = "돋음체"
        .RowHeightMin = 500
        
        .ColFormat(2) = "#,##0"
        .ColFormat(3) = "#,##0"
        .ColFormat(4) = "#,##0"
        .ColFormat(5) = "#,##0"
        


        .ExplorerBar = flexExNone
        .Editable = flexEDNone
        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
    
    With grdOrder
        .Cols = 7
        Call SetVSFlexGrid(grdOrder)

        .Redraw = flexRDNone

        .Rows = 5
        .FixedRows = 5
        
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .RowHidden(3) = True

        .TextMatrix(4, 0) = "":                   .ColWidth(0) = 300:         .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(4, 1) = "OrderID":            .ColWidth(1) = 1300:        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(4, 2) = "색상명":             .ColWidth(2) = 2300:        .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(4, 3) = "LotNO":              .ColWidth(3) = 800:         .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(4, 4) = "절수":               .ColWidth(4) = 800:         .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(4, 5) = "수량(Y)":            .ColWidth(5) = 1200:        .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(4, 6) = "수량(M)":            .ColWidth(6) = 1200:        .ColAlignment(6) = flexAlignRightCenter
        
        .ExplorerBar = flexExNone
        .Editable = flexEDNone
        .ExtendLastCol = True
        .Redraw = flexRDDirect
    End With
    

End Sub

Private Sub FillGridOrder()
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    Dim sDate$, sCustomID$, sArticleID$, sWorkID$, nUnitPrice As Integer
    Dim nQtyYDS As Long, nQtyMET As Long
    
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlSub.Visible = True
        
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    With grdData
        sDate = MakeDate(DF_SHORT, dtpDate)
        sCustomID = .TextMatrix(.Row, 9)
        sArticleID = .TextMatrix(.Row, 10)
        sWorkID = .TextMatrix(.Row, 11)
        nUnitPrice = .ValueMatrix(.Row, 6)
    End With
    
    nQtyMET = 0: nQtyYDS = 0
    
    With grdOrder
        .Rows = .FixedRows
        .Redraw = flexRDNone
    
        Set rs = oOrder.GetResultSaleExpectDetail(sDate, sCustomID, sArticleID, sWorkID, nUnitPrice)
        
        Do Until rs.EOF
            .AddItem "" & vbTab & MakeOrderID(rs!OrderID, OM_EXPAND) & vbTab & rs!Color & vbTab & rs!LotNo & vbTab & _
                         rs!RollNo & vbTab & IIf(rs!CtrlQtyYDS = 0, "", Format(rs!CtrlQtyYDS, "#,##0")) & vbTab & IIf(rs!CtrlQtyMET = 0, "", Format(rs!CtrlQtyMET, "#,##0"))
            nQtyYDS = nQtyYDS + rs!CtrlQtyYDS
            nQtyMET = nQtyMET + rs!CtrlQtyMET
            rs.MoveNext
        Loop

        rs.Close
        Set rs = Nothing
        
        If nQtyMET <> 0 Then
            nQtyYDS = nQtyYDS + Int(nQtyMET / 0.9144)
        End If
        
        .AddItem "" & vbTab & "" & vbTab & "합    계" & vbTab & "" & vbTab & _
                         "" & vbTab & Format(nQtyYDS, "#,##0") & vbTab & ""
        
        .Redraw = flexRDDirect
    End With
    
    pnlProgress.Visible = False
    
    Exit Sub

ErrHandler:
    pnlProgress.Visible = False
    Set oOrder = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmResultSaleExpect.FillGridOrder", Err.Description)

End Sub
Private Sub FillGridData()
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    Dim i%, sSDate$, sEDate$
    
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    sSDate = MakeDate(DF_SHORT, dtpDate)
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    With grdData
        .Rows = .FixedRows
        .Redraw = flexRDNone
    
        Set rs = oOrder.GetResultSaleExpect(sSDate)
        
        Do Until rs.EOF
            .AddItem .Rows - .FixedRows + 1 & vbTab & rs!kCustom & vbTab & rs!Article & vbTab & rs!WorkName & vbTab & _
                         rs!CtrlQty & vbTab & 0 & vbTab & rs!UnitPrice & vbTab & Format(rs!Price, "#,##0") & vbTab & "" & vbTab & _
                         rs!CustomID & vbTab & rs!ArticleID & vbTab & rs!WorkID & vbTab & rs!sLevel
            If rs!sLevel = "2" Or rs!sLevel = "3" Then
'                .TextMatrix(.Rows - 1, 1) = rs!Article
                .TextMatrix(.Rows - 1, 2) = ""
'                .Cell(flexcpBackColor, .Rows - 1, 2, .Rows - 1, 2) = PRNHeaderColor
'                .Cell(flexcpBackColor, .Rows - 1, 5, .Rows - 1, 5) = PRNHeaderColor
                .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, 1) = flexAlignCenterCenter
                
                .Cell(flexcpFontBold, .Rows - 1, 2, .Rows - 1, .Cols - 1) = True
                .Cell(flexcpFontBold, .Rows - 1, 5, .Rows - 1, .Cols - 1) = True
                .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
                
                .TextMatrix(.Rows - 1, 0) = ""
                
            End If
            rs.MoveNext
        Loop

        If rs.RecordCount > 0 Then
            .Row = .FixedRows
            .Col = 1
        End If
        rs.Close
        Set rs = Nothing
        
        .Redraw = flexRDDirect
    End With
    
    pnlProgress.Visible = False
    
    Exit Sub

ErrHandler:
    pnlProgress.Visible = False
    Set oOrder = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmResultSaleExpect.FillGridData", Err.Description)
End Sub

'Private Sub grdData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
''    With grdData
''        .TextMatrix(12, 2) = CInt((CheckNum(.TextMatrix(12, 1)) * 100) / 62)
''    End With
'End Sub

'Private Sub grdData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
''    If Row < 12 Or Col > 1 Then Cancel = True
'End Sub

Private Sub grdDataSelect()
    With grdData
        If .TextMatrix(.Row, .Cols - 1) = "1" Then
            If .TextMatrix(.Row, 0) < (.TopRow + 3) Then
                pnlSub.Top = 4700
            Else
                pnlSub.Top = 960
            End If
            
            Call FillGridOrder
        Else
            pnlSub.Visible = False
            
        End If
    End With

End Sub

Private Sub grdData_RowColChange()

    If optView(0).Value = True Then
        Call grdDataSelect
    Else
        pnlSub.Visible = False

    End If

End Sub

Private Sub optView_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
        pnlSub.Visible = True
    Else
        pnlSub.Visible = False
    End If

End Sub
