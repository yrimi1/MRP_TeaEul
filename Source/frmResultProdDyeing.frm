VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmResultProdDyeing 
   Caption         =   "월 생산 집계 및 사고율 현황"
   ClientHeight    =   9315
   ClientLeft      =   1020
   ClientTop       =   2625
   ClientWidth     =   15240
   Icon            =   "frmResultProdDyeing.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   15240
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9060
      MouseIcon       =   "frmResultProdDyeing.frx":000C
      MousePointer    =   2  '십자형
      ScaleHeight     =   13
      ScaleMode       =   0  '사용자
      ScaleWidth      =   11
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Frame fraSearch 
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   -90
      Width           =   5880
      Begin VB.ComboBox cboYear 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "2002"
         Top             =   120
         Width           =   1125
      End
      Begin VB.ComboBox cboMonth 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2475
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "01"
         Top             =   120
         Width           =   795
      End
      Begin Threed.SSPanel pnlName 
         Height          =   405
         Index           =   4
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   714
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "년월 선택"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   420
         Left            =   4650
         TabIndex        =   6
         Tag             =   "PERM_ADDNEW"
         Top             =   120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   741
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
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
      Begin Threed.SSCommand cmdTerm 
         Height          =   420
         Index           =   0
         Left            =   3300
         TabIndex        =   9
         Tag             =   "PERM_ADDNEW"
         Top             =   120
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   741
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
         Caption         =   "금월"
         PictureAlignment=   1
      End
      Begin Threed.SSCommand cmdTerm 
         Height          =   420
         Index           =   1
         Left            =   3900
         TabIndex        =   10
         Tag             =   "PERM_ADDNEW"
         Top             =   120
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   741
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
         Caption         =   "전월"
         PictureAlignment=   1
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   450
      Left            =   13980
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   794
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "         닫기"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdResult 
      Height          =   8805
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   15210
      _cx             =   26829
      _cy             =   15531
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   100
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin Threed.SSCommand cmdPrint 
      Height          =   450
      Left            =   5880
      TabIndex        =   8
      Tag             =   "PERM_ADDNEW"
      Top             =   0
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   794
      _Version        =   196609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "        발행"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmResultProdDyeing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
Dim iCol%

    With grdResult
        .Redraw = flexRDNone
        
        .RowHeight(0) = 700
        .RowHeight(1) = 400
        .ExtendLastCol = True


        .TextMatrix(3, 0) = "구분":   .ColWidth(0) = 1000
        
        .ColWidth(0) = 600
        For iCol = 1 To 31
            .ColWidth(iCol) = 460
        Next iCol
'        .ColWidth(32) = 600

        .FontSize = 5

        .Cell(flexcpText, 0, 1, 0, .Cols - 1) = cboYear.Text & "년 " & cboMonth.Text & "월 " & "생산 집계 및 사고율 현황"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpFontSize, 1, 0, 1, .Cols - 1) = 9
        .Cell(flexcpText, 1, 25, 1, .Cols - 1) = "▶ 발행일 : " & Format(Now, "YYYY년 MM월 DD일  hh시 mm분")
        .Cell(flexcpBackColor, 0, 0, 1, .Cols - 1) = vbWhite

        .PrintGrid "태을염직", True, 2, 50, 800
    
        .ExtendLastCol = False
    
        .ColWidth(0) = 1000
        For iCol = 1 To 31
            .ColWidth(iCol) = 800
        Next iCol
        .ColWidth(32) = 1000
    
        .FontSize = 9
        .RowHeight(0) = 0
        .RowHeight(1) = 0
    
        .Redraw = flexRDDirect
        MsgBox "월 생산집계가 인쇄되었습니다", vbInformation + vbOKOnly, "인쇄 완료"
    End With
End Sub

Private Sub cmdSearch_Click()
    Call InitGrid
    Call InitGraph
    Call FillGridResult
    Call GraphAndSumUp
End Sub

Private Sub FillGridResult()
Dim oRapid As PlusLib2.CRapid
Dim rs As Recordset
Dim iCount%, iCol%

On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
    
    Set oRapid = New PlusLib2.CRapid
    oRapid.Connection = g_adoCon

    With grdResult
        .Redraw = flexRDNone

        Set rs = oRapid.GetResultProdDyeTrouble(cboYear.Text, cboMonth.Text)
        Set oRapid = Nothing
        
        If rs.RecordCount > 0 Then
        
            For iCount = 1 To rs.RecordCount
                iCol = CInt(rs!DayQry)
                
                .TextMatrix(5, iCol) = Format(rs!ProdQty, "##,###")
                .TextMatrix(7, iCol) = Format(rs!DyeQty, "##,###")
                .TextMatrix(9, iCol) = Format(rs!ManuQty, "##,###")
                .TextMatrix(11, iCol) = Format(rs!DyeQty + rs!ManuQty, "##,###")
                If rs!ProdQty = 0 Then
                    .TextMatrix(13, iCol) = ""
                Else
                    .TextMatrix(13, iCol) = Format((rs!DyeQty + rs!ManuQty) / rs!ProdQty * 100, "##0.0")
                End If
                .TextMatrix(15, iCol) = ""
                .TextMatrix(17, iCol) = Format(rs!InstQty, "##,###")
                .TextMatrix(19, iCol) = Format(rs!ReWorkQty + rs!StainQty + rs!DeColorQty + rs!ModiQty + rs!EtcQty, "##,###")
                .TextMatrix(21, iCol) = Format(rs!ReWorkQty, "##,###")
                .TextMatrix(23, iCol) = Format(rs!StainQty, "##,###")
                .TextMatrix(25, iCol) = Format(rs!DeColorQty, "##,###")
                .TextMatrix(27, iCol) = Format(rs!ModiQty, "##,###")
                .TextMatrix(29, iCol) = Format(rs!EtcQty, "##,###")
                If rs!InstQty = 0 Then
                    .TextMatrix(31, iCol) = ""
                Else
                    .TextMatrix(31, iCol) = Format((rs!ReWorkQty + rs!StainQty + rs!DeColorQty + rs!ModiQty + rs!EtcQty) / rs!InstQty * 100, "##0.0")
                End If
                .TextMatrix(35, iCol) = Format(rs!WeavQty, "##,###")
                .TextMatrix(37, iCol) = ""     ' 출고량
                
                rs.MoveNext
            Next iCount
        End If
        
        Set rs = Nothing
        
        .Redraw = flexRDDirect
    End With

    Screen.MousePointer = vbDefault

    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oRapid = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorBox(Err.Number, "frmResultProdDyeing.FillGridResult", Err.Description)
End Sub

Private Sub GraphAndSumUp()
Dim iCol%
Dim xT1%, yT1%, xT2%, yT2%  ' 사고율 관련 좌표
Dim xD1%, yD1%, xD2%, yD2%  ' 수정율 관련 좌표
Dim nSumProd&, nSumDye&, nSumManu&, nSumTrouble&
Dim nSumInst&, nSumTotModi&, nSumReWork&, nSumStain&
Dim nSumDeColor&, nSumModi&, nSumETC&, nSumWeav&, nSumOut&

    With grdResult
        .Redraw = flexRDNone
            xT1 = 0:    xT2 = 0:    yT1 = 225:  yT2 = 255
            xD1 = 0:    xD2 = 0:    yD1 = 225:  yD2 = 255
            ' 13행(사고율), 31행(수정율)
            picGraph.DrawStyle = vbSolid
            
            For iCol = 1 To 31
                xT2 = 30 * iCol
                If Trim(.TextMatrix(13, iCol)) = "" Then
                    yT2 = 225
                Else
                    yT2 = 225 - CInt(CSng(.TextMatrix(13, iCol)) * 5)
                End If
                picGraph.DrawWidth = 1
                picGraph.Line (xT1, yT1)-(xT2, yT2), RGB(255, 0, 0)
                xT1 = xT2
                yT1 = yT2
                
                xD2 = 30 * iCol
                If Trim(.TextMatrix(31, iCol)) = "" Then
                    yD2 = 225
                Else
                    yD2 = 225 - CInt(CSng(.TextMatrix(31, iCol)) * 5)
                End If
                picGraph.DrawWidth = 2
                picGraph.Line (xD1, yD1)-(xD2, yD2), RGB(0, 0, 255)
                xD1 = xD2
                yD1 = yD2
                
            
                If Trim(.TextMatrix(5, iCol)) <> "" Then
                    nSumProd = nSumProd + CLng(.TextMatrix(5, iCol))
                End If
                If Trim(.TextMatrix(7, iCol)) <> "" Then
                    nSumDye = nSumDye + CLng(.TextMatrix(7, iCol))
                End If
                If Trim(.TextMatrix(9, iCol)) <> "" Then
                    nSumManu = nSumManu + CLng(.TextMatrix(9, iCol))
                End If
                If Trim(.TextMatrix(11, iCol)) <> "" Then
                    nSumTrouble = nSumTrouble + CLng(.TextMatrix(11, iCol))
                End If
                
                If Trim(.TextMatrix(17, iCol)) <> "" Then
                    nSumInst = nSumInst + CLng(.TextMatrix(17, iCol))
                End If
                If Trim(.TextMatrix(19, iCol)) <> "" Then
                    nSumTotModi = nSumTotModi + CLng(.TextMatrix(19, iCol))
                End If
                If Trim(.TextMatrix(21, iCol)) <> "" Then
                    nSumReWork = nSumReWork + CLng(.TextMatrix(21, iCol))
                End If
                If Trim(.TextMatrix(23, iCol)) <> "" Then
                    nSumStain = nSumStain + CLng(.TextMatrix(23, iCol))
                End If
                If Trim(.TextMatrix(25, iCol)) <> "" Then
                    nSumDeColor = nSumDeColor + CLng(.TextMatrix(25, iCol))
                End If
                If Trim(.TextMatrix(27, iCol)) <> "" Then
                    nSumModi = nSumModi + CLng(.TextMatrix(27, iCol))
                End If
                If Trim(.TextMatrix(29, iCol)) <> "" Then
                    nSumETC = nSumETC + CLng(.TextMatrix(29, iCol))
                End If
                If Trim(.TextMatrix(35, iCol)) <> "" Then
                    nSumWeav = nSumWeav + CLng(.TextMatrix(35, iCol))
                End If
                
            Next iCol
            
            .TextMatrix(5, 32) = Format(nSumProd, "###,###")
            .TextMatrix(7, 32) = Format(nSumDye, "###,###")
            .TextMatrix(9, 32) = Format(nSumManu, "###,###")
            .TextMatrix(11, 32) = Format(nSumTrouble, "###,###")
            If nSumProd = 0 Then
                .TextMatrix(13, 32) = ""
            Else
                .TextMatrix(13, 32) = Format(nSumTrouble / nSumProd * 100, "##0.0") & " %"
            End If
            .TextMatrix(17, 32) = Format(nSumInst, "###,###")
            .TextMatrix(19, 32) = Format(nSumTotModi, "###,###")
            .TextMatrix(21, 32) = Format(nSumReWork, "###,###")
            .TextMatrix(23, 32) = Format(nSumStain, "###,###")
            .TextMatrix(25, 32) = Format(nSumDeColor, "###,###")
            .TextMatrix(27, 32) = Format(nSumModi, "###,###")
            .TextMatrix(29, 32) = Format(nSumETC, "###,###")
            If nSumInst = 0 Then
                .TextMatrix(31, 32) = ""
            Else
                .TextMatrix(31, 32) = Format(nSumTotModi / nSumInst * 100, "##0.0") & " %"
            End If
            
            .TextMatrix(35, 32) = Format(nSumWeav, "###,###")
'            .TextMatrix(37, 32) = Format(nSumOut, "###,###")
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then
        cboYear.Text = Left(Format(Now, "YYYYMM"), 4)
        cboMonth.Text = Right(Format(Now, "YYYYMM"), 2)
    Else
        cboYear.Text = Left(Format(DateSerial(Year(Now), Month(Now) - 1, Day(Now)), "YYYYMM"), 4)
        cboMonth.Text = Right(Format(DateSerial(Year(Now), Month(Now) - 1, Day(Now)), "YYYYMM"), 2)
    End If
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = False
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 15300, 9660
    
    Call SetOperate(Me)
    Call CboListAdd
    Call InitGrid
    Call InitGraph
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CboListAdd()
    Dim iCount As Integer
    
    With cboYear
        .Clear
        For iCount = 1 To 3
            .AddItem Year(Now) - iCount
            .AddItem Year(Now) + iCount
        Next iCount
        .AddItem Year(Now)
        .Text = Year(Now)
    End With

    With cboMonth
        .Clear
        For iCount = 1 To 12
            .AddItem Format(iCount, "00")
        Next iCount
        .Text = Format(Month(Now), "00")
    End With
End Sub

Private Sub InitGrid()
    Dim iCol%, irow%, iDay%
    Dim sDate$
    Dim dDate As Date
    
    Call SetVSFlexGrid(grdResult)
    With grdResult
        .Redraw = flexRDNone
        .FontSize = 9
        .WordWrap = False
        .ScrollBars = flexScrollBarHorizontal
        .HighLight = flexHighlightNever
        .ExtendLastCol = False
        .Rows = 39:         .Cols = 34
        .FixedRows = 4:     .FixedCols = 1
        
        .RowHeightMin = 0
        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .RowHeight(2) = 0
        
        For iCol = 0 To .Cols - 1
            .ColWidth(iCol) = 0
            .ColAlignment(iCol) = flexAlignRightCenter
        Next iCol
        .ColAlignment(0) = flexAlignCenterCenter

        For irow = 0 To .Rows - 1
            .RowHeight(irow) = 0
        Next irow
        
        .RowHeight(3) = 300
        .TextMatrix(3, 0) = "구분":   .ColWidth(0) = 1000
        
        dDate = CDate(cboYear & "-" & cboMonth & "-" & "01")
        
        dDate = DateSerial(Year(dDate), Month(dDate) + 1, 1 - 1)
        For iCol = 1 To 31
            If CInt(Format(dDate, "DD")) >= iCol Then
                sDate = cboYear & "-" & cboMonth & "-" & Format(iCol, "00")
                If Weekday(CDate(sDate)) = 1 Then  ' 일요일
                    .Cell(flexcpForeColor, 3, iCol) = vbRed
                End If
            End If
            iDay = iDay + 1
            .TextMatrix(3, iCol) = CStr(iDay):      .ColWidth(iCol) = 800
        Next iCol
        .TextMatrix(3, 32) = "TOTAL":               .ColWidth(32) = 1000
        .TextMatrix(3, 33) = "%":                   .ColWidth(33) = 0
        
        .TextMatrix(5, 0) = "생산량":       .RowHeight(5) = 300
        .TextMatrix(7, 0) = "염색사고":     .RowHeight(7) = 300
        .TextMatrix(9, 0) = "가공사고":     .RowHeight(9) = 300
        .TextMatrix(11, 0) = "사고량":      .RowHeight(11) = 300
        .TextMatrix(13, 0) = "%":           .RowHeight(13) = 300
        .TextMatrix(15, 0) = "":            .RowHeight(15) = 30
        .Cell(flexcpBackColor, 15, 0, 15, .Cols - 1) = vbBlue
        .TextMatrix(17, 0) = "염색투입":    .RowHeight(17) = 300
        .TextMatrix(19, 0) = "수정량":      .RowHeight(19) = 300
        .TextMatrix(21, 0) = "색수정,재염": .RowHeight(21) = 300
        .TextMatrix(23, 0) = "얼룩,오염":   .RowHeight(23) = 300
        .TextMatrix(25, 0) = "탈발,탈색":   .RowHeight(25) = 300
        .TextMatrix(27, 0) = "시와,수정":   .RowHeight(27) = 300
        .TextMatrix(29, 0) = "기타":        .RowHeight(29) = 300
        .TextMatrix(31, 0) = "%":           .RowHeight(31) = 300
        .TextMatrix(33, 0) = "":            .RowHeight(33) = 3700
        .Cell(flexcpAlignment, 33, 0) = flexAlignRightTop
        .TextMatrix(35, 0) = "제직불량":    .RowHeight(35) = 300
        .TextMatrix(37, 0) = "출고량":      .RowHeight(37) = 300
        
        .Cell(flexcpText, 33, 1, 33, .Cols - 3) = " "
        .MergeCells = flexMergeRestrictRows
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(33) = True
        
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub InitGraph()
Dim iXPos%, iYPos%
'       (X1, Y1) - (X2, Y2)
'-------------------------------------------------
' 45%   (0, 25) - (1000, 25)        Y위치: 25
' 40%   (0, 50) - (1000, 50)        Y위치: 50
' ..
' ..
' 0%    (0, 225) - (1000, 225)      Y위치: 225
    With picGraph
        .Visible = False
        grdResult.Col = 1
        .Left = grdResult.CellLeft
        .Width = 15000
        .Height = grdResult.RowHeight(33)
        .Cls
        
        picGraph.DrawWidth = 1
    
        .DrawStyle = vbDot
        For iYPos = 25 To 200 Step 25
            picGraph.Line (0, iYPos)-(1000, iYPos), &H8000000F
        Next iYPos
        
        .CurrentX = 3
        .CurrentY = 20
        picGraph.Print "40"
        .CurrentX = 3
        .CurrentY = 45
        picGraph.Print "35"
        .CurrentX = 3
        .CurrentY = 70
        picGraph.Print "30"
        .CurrentX = 3
        .CurrentY = 95
        picGraph.Print "25"
        .CurrentX = 3
        .CurrentY = 120
        picGraph.Print "20"
        .CurrentX = 3
        .CurrentY = 145
        picGraph.Print "15"
        .CurrentX = 3
        .CurrentY = 170
        picGraph.Print "10"
        .CurrentX = 3
        .CurrentY = 195
        picGraph.Print "5"
        
        For iXPos = 30 To 930 Step 30
            picGraph.Line (iXPos, 0)-(iXPos, 225), &H8000000F
            .CurrentX = iXPos - 9
            .CurrentY = 230
            picGraph.Print CStr(iXPos \ 30)
        Next iXPos
        
        .CurrentX = 45
        .CurrentY = 10
        picGraph.Print "사고율"
        
        .CurrentX = 45
        .CurrentY = 35
        picGraph.Print "수정율"
        
        .DrawStyle = vbSolid
        picGraph.Line (0, 0)-(0, 225), RGB(0, 0, 0)
        picGraph.Line (0, 225)-(1000, 225), RGB(0, 0, 0)
        
        picGraph.Line (25, 15)-(40, 15), RGB(255, 0, 0)
        picGraph.DrawWidth = 2
        picGraph.Line (25, 40)-(40, 40), RGB(0, 0, 255)
        picGraph.DrawWidth = 1
        
        .Visible = True
        grdResult.Cell(flexcpPicture, 33, 1, 33, grdResult.Cols - 3) = .Image
        .Visible = False
    End With

End Sub
