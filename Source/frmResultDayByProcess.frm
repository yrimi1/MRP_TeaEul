VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmResultDayByProcess 
   ClientHeight    =   9255
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   11850
   Icon            =   "frmResultDayByProcess.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11850
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
      Left            =   8310
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
      Left            =   10080
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
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   11835
      _cx             =   20876
      _cy             =   13309
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
   Begin Threed.SSFrame frmSearch 
      Height          =   825
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
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
         Left            =   8100
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
         Format          =   113639424
         CurrentDate     =   37957
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Left            =   60
         TabIndex        =   4
         Top             =   210
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
         Caption         =   "실적 일자"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   8130
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
         Left            =   10125
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
         Left            =   5235
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
End
Attribute VB_Name = "frmResultDayByProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLeft_Click()
    dtpDate = dtpDate - 1
    Call cmdSearch_Click
End Sub

Private Sub cmdPrint_Click()
    Dim oExcel      As Excel.Application
    Dim oExcelBook  As Excel.Workbook
    Dim oExcelSheet As Excel.Worksheet
    Dim oFs         As FileSystemObject
    Dim i%, j%, sWeek$, sReport$
    
    On Error GoTo ErrHandler
    
    Me.PopupMenu PlusMDI.mnuPopup
    
    Set oExcel = New Excel.Application
    Set oExcelBook = oExcel.Workbooks.Open(App.Path & "\Report\ProcessDayResult.xls")

    With grdData
        Select Case Weekday(dtpDate, vbSunday)
            Case 1:
                sWeek = "일요일"
            Case 2:
                sWeek = "월요일"
            Case 3:
                sWeek = "화요일"
            Case 4:
                sWeek = "수요일"
            Case 5:
                sWeek = "목요일"
            Case 6:
                sWeek = "금요일"
            Case 7:
                sWeek = "토요일"
        End Select
        
        oExcel.Cells(5, 2) = Format(dtpDate, "YYYY년 M월 D일 ") & sWeek
        For i = 1 To .Rows - 1
            If i < 11 Then
                oExcel.Cells(i + 6, 2) = .TextMatrix(i, 1)
                oExcel.Cells(i + 6, 5) = .TextMatrix(i, 3)
            ElseIf i = 11 Then
                oExcel.Cells(i + 6, 2) = Right(.TextMatrix(i, 1), Len(.TextMatrix(i, 1)) - 1)
                oExcel.Cells(i + 6, 5) = Right(.TextMatrix(i, 3), Len(.TextMatrix(i, 3)) - 1)
            ElseIf i = 12 Then
                 oExcel.Cells(i + 6, 2) = .TextMatrix(i, 1)
            End If
        Next i
        
    End With

    sReport = App.Path & "\Report\TmpProcessDayResult.xls"
    Set oFs = New FileSystemObject
    If oFs.FileExists(sReport) Then Call oFs.DeleteFile(sReport)
    Set oFs = Nothing

    Call oExcelBook.SaveAs(sReport)

    If PlusMDI.PrintPreview Then
        oExcel.WindowState = xlMaximized
        oExcel.Application.Visible = True
'        oExcel.ActiveWindow.SelectedSheets.PrintPreview
    Else
        oExcel.ActiveWindow.SelectedSheets.PrintOut Copies:=1
        Call ProcessClose("XLMAIN")
    End If
    
    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing
    
    
    
    Exit Sub
    
ErrHandler:
    Set oExcelSheet = Nothing
    Set oExcelBook = Nothing
    Set oExcel = Nothing

    Call ErrorBox(Err.Number, "frmResultDayByProcess", Err.Description)
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
    
    Call FillGridData
    pnlProgress.Visible = False
End Sub

Private Sub InitGrid()
    Dim i%
    
    With grdData
        .Cols = 5
        Call SetVSFlexGrid(grdData)

        .Redraw = flexRDNone

        .Rows = 1
        .FixedRows = 1

        .TextArray(0) = "구  분":           .ColWidth(0) = 2000:        .ColAlignment(0) = flexAlignCenterCenter
        .TextArray(1) = "생산량(Y)":        .ColWidth(1) = 2500:        .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "가동율(%)":        .ColWidth(2) = 2500:        .ColAlignment(2) = flexAlignCenterCenter
        .TextArray(3) = "생산량 누계(Y)":   .ColWidth(3) = 2500:        .ColAlignment(3) = flexAlignCenterCenter
        .TextArray(4) = "가동율(%)":        .ColWidth(4) = 2500:        .ColAlignment(4) = flexAlignCenterCenter
        
        .AddItem "ORDER 접수"
        .AddItem "입  고"
        .AddItem "정  련"
        .AddItem "수  세"
        .AddItem "C.P.B"
        .AddItem "PEACH"
        .AddItem "염  색"
        .AddItem "가  공"
        .AddItem "검  사"
        .AddItem "출  고"
        .AddItem "매출액"
        .AddItem "출근인원"
        
        .ColFormat(1) = "#,##0"
        .ColFormat(2) = "#,##0"
        .ColFormat(3) = "#,##0"
        .ColFormat(4) = "#,##0"
        
        .RowHeight(0) = 700
        For i = 1 To .Rows - 1
            .RowHeight(i) = 670
        Next i
        
        .RowHidden(4) = True
        .RowHidden(6) = True
        
        .FontSize = 12
        .Cell(flexcpBackColor, 1, 3, 11, 4) = COLOR_GRIDROW

        .Editable = flexEDKbdMouse
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FillGridData()
    Dim oOrder As PlusLib2.COrder
    Dim rs As ADODB.Recordset
    Dim i%, sSDate$, sEDate$
    
    On Error GoTo ErrHandler
    
    proProgress.Value = 0
    lblCount = LoadResString(160)
    pnlProgress.Visible = True
        
    sSDate = Left(MakeDate(DF_SHORT, dtpDate), 6) + "01"
    sEDate = MakeDate(DF_SHORT, dtpDate)
    
    Set oOrder = New PlusLib2.COrder
    oOrder.Connection = g_adoCon
    
    With grdData
        .Redraw = flexRDNone
    
        Set rs = oOrder.GetResultDayByProcess(sEDate, sEDate, val(txtExchange))
        
        .TextMatrix(1, 1) = rs!OrderQty
        .TextMatrix(1, 2) = CInt(rs!OrderQty / 600) ' 6000
        
        .TextMatrix(2, 1) = rs!StuffInQty
        .TextMatrix(2, 2) = CInt(rs!StuffInQty / 400) '40000
        
        .TextMatrix(3, 1) = rs!CPBPreQty
        .TextMatrix(3, 2) = CInt(rs!CPBPreQty / 600) '60000
        
        .TextMatrix(4, 1) = rs!RefineQty
        .TextMatrix(4, 2) = CInt(rs!RefineQty / 1500) '150000
        
        .TextMatrix(5, 1) = rs!CRapidQty
        .TextMatrix(5, 2) = CInt(rs!CRapidQty / 200) '20000
        
        .TextMatrix(6, 1) = rs!PeachQty
        .TextMatrix(6, 2) = CInt(rs!PeachQty / 400) '40000
        
        .TextMatrix(7, 1) = rs!RapidQTy
        .TextMatrix(7, 2) = CInt(rs!RapidQTy / 500) '50000
        
        .TextMatrix(8, 1) = rs!TenterQty
        .TextMatrix(8, 2) = CInt(rs!TenterQty / 600) '60000
        
        .TextMatrix(9, 1) = rs!InspectQty
        .TextMatrix(9, 2) = CInt(rs!InspectQty / 500) '50000
        
        .TextMatrix(10, 1) = rs!OutQty
        .TextMatrix(10, 2) = CInt(rs!OutQty / 500) '50000
        
        .TextMatrix(11, 1) = "W" & Format(rs!OutPrice, "#,##0")
        .TextMatrix(11, 2) = CInt(rs!OutPrice / 270000) '27000000

        rs.Close
        Set rs = Nothing
        
        Set rs = oOrder.GetResultDayByProcess(sSDate, sEDate, val(txtExchange))
        
        .TextMatrix(1, 3) = rs!OrderQty
        .TextMatrix(1, 4) = CInt(rs!OrderQty / 15600) ' 156000
        .TextMatrix(2, 3) = rs!StuffInQty
        .TextMatrix(2, 4) = CInt(rs!StuffInQty / 10400) '1040000
        .TextMatrix(3, 3) = rs!CPBPreQty
        .TextMatrix(3, 4) = CInt(rs!CPBPreQty / 15600) '1560000
        .TextMatrix(4, 3) = rs!RefineQty
        .TextMatrix(4, 4) = CInt(rs!RefineQty / 39000) '3900000
        .TextMatrix(5, 3) = rs!CRapidQty
        .TextMatrix(5, 4) = CInt(rs!CRapidQty / 5200) '520000
        .TextMatrix(6, 3) = rs!PeachQty
        .TextMatrix(6, 4) = CInt(rs!PeachQty / 10400) '1040000
        .TextMatrix(7, 3) = rs!RapidQTy
        .TextMatrix(7, 4) = CInt(rs!RapidQTy / 13000) '1300000
        .TextMatrix(8, 3) = rs!TenterQty
        .TextMatrix(8, 4) = CInt(rs!TenterQty / 15600) '1560000
        .TextMatrix(9, 3) = rs!InspectQty
        .TextMatrix(9, 4) = CInt(rs!InspectQty / 13000) '1300000
        .TextMatrix(10, 3) = rs!OutQty
        .TextMatrix(10, 4) = CInt(rs!OutQty / 13000) '1300000
        .TextMatrix(11, 3) = "W" & Format(rs!OutPrice, "#,##0")
        .TextMatrix(11, 4) = CInt(rs!OutPrice / 7020000) '702000000

        rs.Close
        Set rs = Nothing
        
        
        .Redraw = flexRDDirect
'        .SetFocus
    End With
    
    pnlProgress.Visible = False
    
    Exit Sub

ErrHandler:
    pnlProgress.Visible = False
    Set oOrder = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmResultDayByProcess.FillGridData", Err.Description)
End Sub

Private Sub grdData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdData
        .TextMatrix(12, 2) = CInt((CheckNum(.TextMatrix(12, 1)) * 100) / 62)
    End With
End Sub

Private Sub grdData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row < 12 Or Col > 1 Then Cancel = True
End Sub

