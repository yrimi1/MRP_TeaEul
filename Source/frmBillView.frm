VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBillView 
   Caption         =   "정산서/청구서 조회 및 발행"
   ClientHeight    =   9915
   ClientLeft      =   1005
   ClientTop       =   765
   ClientWidth     =   14295
   Icon            =   "frmBillView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   14295
   Begin VB.ComboBox cobYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1470
      Sorted          =   -1  'True
      TabIndex        =   14
      Text            =   "2002"
      Top             =   30
      Width           =   1155
   End
   Begin VB.ComboBox cobMonth 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2625
      Sorted          =   -1  'True
      TabIndex        =   13
      Text            =   "12"
      Top             =   30
      Width           =   765
   End
   Begin VB.TextBox txtCustomID 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5070
      TabIndex        =   2
      Top             =   60
      Width           =   2505
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   540
      Left            =   10350
      TabIndex        =   0
      Top             =   0
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   953
      _Version        =   196609
      Caption         =   "        닫기(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   540
      Left            =   8160
      TabIndex        =   1
      Tag             =   "PERM_ADDNEW"
      Top             =   0
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   953
      _Version        =   196609
      Caption         =   "        조회(&F)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdFind 
      Height          =   375
      Left            =   7590
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   30
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   661
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
   Begin Threed.SSPanel pnlCaption 
      Height          =   345
      Index           =   6
      Left            =   3660
      TabIndex        =   4
      Top             =   60
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   609
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkCustom 
         Caption         =   "거래처"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   60
         Width           =   990
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8805
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   15531
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   706
      TabMaxWidth     =   7056
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "정산서"
      TabPicture(0)   =   "frmBillView.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdBill"
      Tab(0).Control(1)=   "pnlWaitTab(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "청구서"
      TabPicture(1)   =   "frmBillView.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grdBillAccount"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "pnlWaitTab(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VSFlex7LCtl.VSFlexGrid grdBill 
         Height          =   8235
         Left            =   -74940
         TabIndex        =   7
         Top             =   480
         Width           =   11745
         _cx             =   20717
         _cy             =   14526
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
         BackColorSel    =   16777215
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
      Begin Threed.SSPanel pnlWaitTab 
         Height          =   375
         Index           =   0
         Left            =   -74910
         TabIndex        =   8
         Top             =   60
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   5
         ForeColor       =   16777215
         BackColor       =   14389120
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "정산서"
         BevelOuter      =   0
         FloodColor      =   14389120
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdPrint1 
            Height          =   390
            Left            =   2460
            TabIndex        =   11
            Tag             =   "PERM_ADDNEW"
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   688
            _Version        =   196609
            Caption         =   "정산서 발행"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel pnlWaitTab 
         Height          =   375
         Index           =   1
         Left            =   4140
         TabIndex        =   9
         Top             =   60
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   5
         ForeColor       =   16777215
         BackColor       =   14389120
         PictureMaskColor=   16777215
         MarqueeDelay    =   700
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "청구서"
         BevelOuter      =   0
         FloodColor      =   14389120
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdPrint2 
            Height          =   390
            Left            =   2490
            TabIndex        =   12
            Tag             =   "PERM_ADDNEW"
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   688
            _Version        =   196609
            Caption         =   "청구서 발행"
            PictureAlignment=   1
            RoundedCorners  =   0   'False
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grdBillAccount 
         Height          =   8235
         Left            =   60
         TabIndex        =   10
         Top             =   480
         Width           =   11745
         _cx             =   20717
         _cy             =   14526
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
         BackColorSel    =   16777215
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
   End
   Begin VB.Label lblMainTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "■ 정산 월"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   300
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmBillView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCustom_Click()
    If chkCustom.Value = vbChecked Then
        txtCustomID.Enabled = True
        cmdFind.Enabled = True
        txtCustomID.SetFocus
    Else
        txtCustomID.Enabled = False
        cmdFind.Enabled = False
        cmdSearch.SetFocus
    End If
End Sub

Private Sub cmdFind_Click()
    Call ReturnCode(LG_CUSTOM, 0, False, txtCustomID)
End Sub

Private Sub cmdPrint1_Click()
Dim sEndDay$
Dim lWidth As Long

    With grdBill
        .Redraw = False
        
        lWidth = .Width
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = vbWhite
        .RowHeight(0) = 500
        .RowHeight(2) = 350
        .FontSize = 8
        .Width = lWidth - 800
        
        .ColWidth(2) = 2200
        .ColWidth(4) = 850
        sEndDay = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(cobYear & "-" & cobMonth & "-" & "01"))), "YYYYMMDD")
    
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = "수 불  명 세 서"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 2, .Cols - 1) = True
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpText, 2, 1, 2, 1) = "[[ 태을염직 ]]"
        .Cell(flexcpText, 2, 7, 2, 9) = "▶ 정산기간 : " & cobYear & "년" & cobMonth & "월" & "01일" & " ~ " & cobYear & "년" & cobMonth & "월" & Right(sEndDay, 2) & "일"
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(2) = True
        
        .PrintGrid "태을염직", True, 1, 100, 700
        .FontSize = 9
        .ColWidth(2) = 2800
        .ColWidth(4) = 1000
        .Width = lWidth
        .GridLinesFixed = flexGridInset
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = 0
        .GridColorFixed = &H80000010
        .RowHeight(0) = 0
        .RowHeight(2) = 0
        
        
        .Redraw = True
        
    End With
End Sub

Private Sub cmdPrint2_Click()
Dim sEndDay$
Dim lWidth As Long

    With grdBillAccount
        .Redraw = False
        
        lWidth = .Width
        .GridLinesFixed = flexGridNone
        .GridColorFixed = vbWhite
        .RowHeight(0) = 500
        .RowHeight(2) = 350
        .FontSize = 8
        .Width = lWidth - 700
        
        .ColWidth(1) = 900
        .ColWidth(2) = 2000
        .ColWidth(3) = 800
        .ColWidth(6) = 600
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
        
        sEndDay = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(cobYear & "-" & cobMonth & "-" & "01"))), "YYYYMMDD")
    
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = "ORDER별   청 구 서"
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 16
        .Cell(flexcpFontBold, 0, 0, 2, .Cols - 1) = True
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        .Cell(flexcpFontUnderline, 0, 0, 0, .Cols - 1) = True
        .Cell(flexcpText, 2, 1, 2, 2) = "[[ 태을염직 ]]"
        .Cell(flexcpText, 2, 6, 2, 10) = "▶ 정산일자 : " & cobYear & "년" & cobMonth & "월" & "01일" & " ~ " & cobYear & "년" & cobMonth & "월" & Right(sEndDay, 2) & "일"
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(2) = True
        
        .PrintGrid "태을염직", True, 1, 100, 700
        .FontSize = 9
        
        .ColWidth(1) = 1200
        .ColWidth(2) = 2200
        .ColWidth(3) = 900
        .ColWidth(6) = 800
        .ColWidth(7) = 1100
        .ColWidth(8) = 1100
        .ColWidth(9) = 1100
        .ColWidth(10) = 1100
        
        .Width = lWidth
        .GridLinesFixed = flexGridInset
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = 0
        .RowHeight(0) = 0
        .RowHeight(2) = 0
        
        
        .Redraw = True
        
    End With

End Sub

Private Sub cmdSearch_Click()
    Call FillGridBill
End Sub

Private Sub Form_Activate()
    If PlusMDI.txtName = "9020" Then
        SSTab1.Tab = 1
    Else
        SSTab1.Tab = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11970, 9660
    
    Call SetOperate(Me)
    cmdFind.Picture = LoadResPicture("FIND", vbResIcon)
    cmdFind.Enabled = False
    txtCustomID.Enabled = False
    
    Call CobListAdd
    Call InitGrid
End Sub

Private Sub CobListAdd()
    Dim iCount As Integer
    
    With cobYear
        .Clear
        For iCount = 1 To 2
            .AddItem Year(Now) - iCount
            .AddItem Year(Now) + iCount
        Next iCount
        .AddItem Year(Now)
        .Text = Year(Now)
    End With

    With cobMonth
        .Clear
        For iCount = 1 To 12
            .AddItem Format(iCount, "00")
        Next iCount
        .Text = Format(Month(Now), "00")
    End With
End Sub

Private Sub InitGrid()
    Dim i%, idx%
    
        With grdBill
            .Redraw = flexRDNone
            
            .SelectionMode = flexSelectionFree
            .ScrollBars = flexScrollBarVertical
'            .ExtendLastCol = False
            
            .Rows = 4:          .Cols = 14
            .FixedRows = 4:     .FixedCols = 0
            
            .RowHeight(0) = 0
            .RowHeight(1) = 0
            .RowHeight(2) = 0
            .RowHeight(3) = 400
    
            For i = 0 To .Cols - 1
                .FixedAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 0
            Next i
    
            .TextMatrix(3, 0) = "구분":         .ColWidth(0) = 0:           .ColAlignment(0) = flexAlignCenterCenter
            .TextMatrix(3, 1) = "거래처":       .ColWidth(1) = 1200:        .ColAlignment(1) = flexAlignLeftCenter
            .TextMatrix(3, 2) = "품명":         .ColWidth(2) = 2800:        .ColAlignment(2) = flexAlignLeftCenter
            .TextMatrix(3, 3) = "ORDER NO":     .ColWidth(3) = 0:           .ColAlignment(3) = flexAlignLeftCenter
            .TextMatrix(3, 4) = "일자":         .ColWidth(4) = 1000:        .ColAlignment(4) = flexAlignCenterCenter
            .TextMatrix(3, 5) = "전월이월":     .ColWidth(5) = 1300:        .ColAlignment(5) = flexAlignRightCenter
            .TextMatrix(3, 6) = "입고수량":     .ColWidth(6) = 1300:        .ColAlignment(6) = flexAlignRightCenter
            .TextMatrix(3, 7) = "출고수량":     .ColWidth(7) = 1300:        .ColAlignment(7) = flexAlignRightCenter
            .TextMatrix(3, 8) = "소요량":       .ColWidth(8) = 1300:        .ColAlignment(8) = flexAlignRightCenter
            .TextMatrix(3, 9) = "재고량":       .ColWidth(9) = 1300:        .ColAlignment(9) = flexAlignRightCenter
            .TextMatrix(3, 10) = "거래처코드":  .ColWidth(10) = 0:          .ColAlignment(10) = flexAlignLeftCenter
            .TextMatrix(3, 11) = "품명코드":    .ColWidth(11) = 0:          .ColAlignment(11) = flexAlignLeftCenter
            .TextMatrix(3, 12) = "orderno":     .ColWidth(12) = 0:          .ColAlignment(12) = flexAlignLeftCenter
            .TextMatrix(3, 13) = "수불일자":    .ColWidth(13) = 0:          .ColAlignment(13) = flexAlignLeftCenter
            
            .Redraw = flexRDDirect
        End With
        

        With grdBillAccount
            .Redraw = flexRDNone
            
            .SelectionMode = flexSelectionFree
            .ScrollBars = flexScrollBarVertical
            
            .Rows = 4:          .Cols = 24
            .FixedRows = 4:     .FixedCols = 0
            
            .RowHeight(0) = 0
            .RowHeight(1) = 0
            .RowHeight(2) = 0
            .RowHeight(3) = 400
    
            For i = 0 To .Cols - 1
                .FixedAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 0
            Next i
    
            .TextMatrix(3, 0) = "구분":         .ColWidth(0) = 0:           .ColAlignment(0) = flexAlignCenterCenter
            .TextMatrix(3, 1) = "거래처":       .ColWidth(1) = 1200:        .ColAlignment(1) = flexAlignLeftCenter
            .TextMatrix(3, 2) = "품명":         .ColWidth(2) = 2200:        .ColAlignment(2) = flexAlignLeftCenter
            .TextMatrix(3, 3) = "출고가공":     .ColWidth(3) = 900:        .ColAlignment(3) = flexAlignLeftCenter
            .TextMatrix(3, 4) = "Order NO.":    .ColWidth(4) = 1200:        .ColAlignment(4) = flexAlignLeftCenter
            .TextMatrix(3, 5) = "ORDER량":      .ColWidth(5) = 900:        .ColAlignment(5) = flexAlignRightCenter
            .TextMatrix(3, 6) = "단가":         .ColWidth(6) = 800:        .ColAlignment(6) = flexAlignRightCenter
            .TextMatrix(3, 7) = "전월누계":     .ColWidth(7) = 1100:        .ColAlignment(7) = flexAlignRightCenter
            .TextMatrix(3, 8) = "금월출고":     .ColWidth(8) = 1100:        .ColAlignment(8) = flexAlignRightCenter
            .TextMatrix(3, 9) = "공급가액":     .ColWidth(9) = 1100:        .ColAlignment(9) = flexAlignRightCenter
            .TextMatrix(3, 10) = "부가세":      .ColWidth(10) = 1100:        .ColAlignment(10) = flexAlignRightCenter
            .TextMatrix(3, 21) = "거래처코드":  .ColWidth(21) = 0:          .ColAlignment(21) = flexAlignLeftCenter
            .TextMatrix(3, 22) = "품명코드":    .ColWidth(22) = 0:          .ColAlignment(22) = flexAlignLeftCenter
            .TextMatrix(3, 23) = "가공코드":    .ColWidth(23) = 0:          .ColAlignment(23) = flexAlignLeftCenter
            
            .Redraw = flexRDDirect
        End With
        
End Sub

Private Sub FillGridBill()
    Dim oSubul As PlusLib2.CSubul
    Dim rs As ADODB.Recordset
    Dim StartDate$, EndDate$, JWMonth$, StartJW$, EndJW$
    Dim sCustom$, sArticle$
    Dim i%, iCnt%
    Dim SubArticle(1 To 5) As Long
    Dim SubCustom(1 To 5) As Long
    
    On Error GoTo ErrHandler
    
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    StartDate = cobYear & cobMonth & "01"
    EndDate = cobYear & cobMonth & "31"
    
    JWMonth = Format(DateAdd("m", -1, CDate(cobYear & "-" & cobMonth & "-" & "01")), "YYYYMM")
    
    StartJW = JWMonth & "01"
    EndJW = JWMonth & "31"
    
    Set rs = oSubul.GetBillByDate(StartDate, EndDate, JWMonth, chkCustom.Value, txtCustomID.Tag)
    Set oSubul = Nothing
        
    With grdBill
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        If rs.RecordCount > 0 Then
            sCustom = rs!kCustom
            sArticle = rs!Article
        
            For i = 1 To rs.RecordCount
                If sCustom <> rs!kCustom Then
                    If sArticle <> rs!Article Then
                        .Rows = .Rows + 1
                        .RowHeight(.Rows - 1) = 300

                        .TextMatrix(.Rows - 1, 0) = ""
                        .TextMatrix(.Rows - 1, 1) = sCustom
                        .TextMatrix(.Rows - 1, 2) = sArticle
                        .Cell(flexcpText, .Rows - 1, 4) = "품명계"
                        .TextMatrix(.Rows - 1, 5) = Format(SubArticle(1), "##,###")
                        .TextMatrix(.Rows - 1, 6) = Format(SubArticle(2), "##,###")
                        .TextMatrix(.Rows - 1, 7) = Format(SubArticle(3), "##,###")
                        .TextMatrix(.Rows - 1, 8) = Format(SubArticle(4), "##,###")
                        .TextMatrix(.Rows - 1, 9) = Format(SubArticle(5), "##,###")
                    End If
                
                    .Rows = .Rows + 1
                    .RowHeight(.Rows - 1) = 300

                    .TextMatrix(.Rows - 1, 0) = ""
                    .TextMatrix(.Rows - 1, 1) = sCustom
                    .Cell(flexcpText, .Rows - 1, 2, .Rows - 1, 2) = "거래처 계"
                    
                    .TextMatrix(.Rows - 1, 5) = Format(SubCustom(1), "##,###")
                    .TextMatrix(.Rows - 1, 6) = Format(SubCustom(2), "##,###")
                    .TextMatrix(.Rows - 1, 7) = Format(SubCustom(3), "##,###")
                    .TextMatrix(.Rows - 1, 8) = Format(SubCustom(4), "##,###")
                    .TextMatrix(.Rows - 1, 9) = Format(SubCustom(5), "##,###")
                    
                    For iCnt = 1 To 5
                        SubArticle(iCnt) = 0
                        SubCustom(iCnt) = 0
                    Next iCnt
                    
                    sCustom = rs!kCustom
                    sArticle = rs!Article
                Else
                    If sArticle <> rs!Article Then
                        .Rows = .Rows + 1
                        .RowHeight(.Rows - 1) = 300

                        .TextMatrix(.Rows - 1, 0) = ""
                        .TextMatrix(.Rows - 1, 1) = sCustom
                        .TextMatrix(.Rows - 1, 2) = sArticle
                        .Cell(flexcpText, .Rows - 1, 4) = "품명계"
                        .TextMatrix(.Rows - 1, 5) = Format(SubArticle(1), "##,###")
                        .TextMatrix(.Rows - 1, 6) = Format(SubArticle(2), "##,###")
                        .TextMatrix(.Rows - 1, 7) = Format(SubArticle(3), "##,###")
                        .TextMatrix(.Rows - 1, 8) = Format(SubArticle(4), "##,###")
                        .TextMatrix(.Rows - 1, 9) = Format(SubArticle(5), "##,###")
                        
                        For iCnt = 1 To 5
                            SubArticle(iCnt) = 0
                        Next iCnt
                        
                        sArticle = rs!Article
                    End If
                End If
                
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 300
                
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = rs!kCustom
                .TextMatrix(.Rows - 1, 2) = rs!Article
                .TextMatrix(.Rows - 1, 3) = rs!OrderNo
                If Len(Trim(rs!subuldate)) = 8 Then
                    .TextMatrix(.Rows - 1, 4) = MakeDate(DF_MD, rs!subuldate)
                    .TextMatrix(.Rows - 1, 5) = ""
                    .TextMatrix(.Rows - 1, 6) = Format(rs!InQty, "##,###")
                    .TextMatrix(.Rows - 1, 7) = Format(rs!OutQty, "##,###")
                    .TextMatrix(.Rows - 1, 8) = Format(rs!syqty, "##,###")
                Else
                    .TextMatrix(.Rows - 1, 4) = ""
                    .TextMatrix(.Rows - 1, 5) = Format(rs!jwqty, "##,###")
                    .TextMatrix(.Rows - 1, 6) = ""
                    .TextMatrix(.Rows - 1, 7) = ""
                    .TextMatrix(.Rows - 1, 8) = ""
                    SubArticle(1) = SubArticle(1) + rs!jwqty
                    SubCustom(1) = SubCustom(1) + rs!jwqty
                    SubArticle(5) = rs!jwqty
                    SubCustom(5) = rs!jwqty
                    
                End If
                SubArticle(5) = SubArticle(5) + rs!InQty - rs!OutQty
                SubCustom(5) = SubCustom(5) + rs!InQty - rs!OutQty
                
'                .TextMatrix(.Rows - 1, 9) = Format(rs!jgqty, "##,###")
                
                .TextMatrix(.Rows - 1, 10) = rs!CustomID
                .TextMatrix(.Rows - 1, 11) = rs!ArticleID
                .TextMatrix(.Rows - 1, 12) = rs!OrderNo
                .TextMatrix(.Rows - 1, 13) = rs!subuldate
                
                SubArticle(2) = SubArticle(2) + rs!InQty
                SubArticle(3) = SubArticle(3) + rs!OutQty
                SubArticle(4) = SubArticle(4) + rs!syqty
                
                SubCustom(2) = SubCustom(2) + rs!InQty
                SubCustom(3) = SubCustom(3) + rs!OutQty
                SubCustom(4) = SubCustom(4) + rs!syqty
                
                rs.MoveNext
            Next i
            
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            .TextMatrix(.Rows - 1, 0) = ""
            .TextMatrix(.Rows - 1, 1) = sCustom
            .TextMatrix(.Rows - 1, 2) = sArticle
            .Cell(flexcpText, .Rows - 1, 4) = "품명계"
            .TextMatrix(.Rows - 1, 5) = Format(SubArticle(1), "##,###")
            .TextMatrix(.Rows - 1, 6) = Format(SubArticle(2), "##,###")
            .TextMatrix(.Rows - 1, 7) = Format(SubArticle(3), "##,###")
            .TextMatrix(.Rows - 1, 8) = Format(SubArticle(4), "##,###")
            .TextMatrix(.Rows - 1, 9) = Format(SubArticle(5), "##,###")
        
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            .TextMatrix(.Rows - 1, 0) = ""
            .TextMatrix(.Rows - 1, 1) = sCustom
            .Cell(flexcpText, .Rows - 1, 2, .Rows - 1, 2) = "거래처 계"
            .TextMatrix(.Rows - 1, 5) = Format(SubCustom(1), "##,###")
            .TextMatrix(.Rows - 1, 6) = Format(SubCustom(2), "##,###")
            .TextMatrix(.Rows - 1, 7) = Format(SubCustom(3), "##,###")
            .TextMatrix(.Rows - 1, 8) = Format(SubCustom(4), "##,###")
            .TextMatrix(.Rows - 1, 9) = Format(SubCustom(5), "##,###")
            
        End If
        rs.Close
        Set rs = Nothing
        
        .MergeCells = flexMergeFree
        .MergeCol(1) = True
        .MergeCol(2) = True
        
        .Redraw = flexRDDirect
        .Row = 0
    End With
    
    
    sCustom = ""
    sArticle = ""
    For iCnt = 1 To 5
        SubArticle(iCnt) = 0
        SubCustom(iCnt) = 0
    Next iCnt
    
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    Set rs = oSubul.GetBillAccountByDate(StartDate, EndDate, StartJW, EndJW, chkCustom.Value, txtCustomID.Tag)
    Set oSubul = Nothing
        
    With grdBillAccount
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
    
        If rs.RecordCount > 0 Then
            sCustom = rs!kCustom
            sArticle = rs!Article
        
            For i = 1 To rs.RecordCount
                If sCustom <> rs!kCustom Then
                    If sArticle <> rs!Article Then
                        .Rows = .Rows + 1
                        .RowHeight(.Rows - 1) = 300

                        .TextMatrix(.Rows - 1, 0) = ""
                        .TextMatrix(.Rows - 1, 1) = sCustom
                        .TextMatrix(.Rows - 1, 2) = sArticle
                        .Cell(flexcpText, .Rows - 1, 3) = "품명 계"
                        .TextMatrix(.Rows - 1, 5) = Format(SubArticle(1), "##,###")
                        .TextMatrix(.Rows - 1, 7) = Format(SubArticle(2), "##,###")
                        .TextMatrix(.Rows - 1, 8) = Format(SubArticle(3), "##,###")
                        .TextMatrix(.Rows - 1, 9) = Format(SubArticle(4), "##,###")
                        .TextMatrix(.Rows - 1, 10) = Format(SubArticle(5), "##,###")
                    End If
                
                    .Rows = .Rows + 1
                    .RowHeight(.Rows - 1) = 300

                    .TextMatrix(.Rows - 1, 0) = ""
                    .TextMatrix(.Rows - 1, 1) = sCustom
                    .Cell(flexcpText, .Rows - 1, 2, .Rows - 1, 2) = "거래처 계"
                    
                    .TextMatrix(.Rows - 1, 5) = Format(SubCustom(1), "##,###")
                    .TextMatrix(.Rows - 1, 7) = Format(SubCustom(2), "##,###")
                    .TextMatrix(.Rows - 1, 8) = Format(SubCustom(3), "##,###")
                    .TextMatrix(.Rows - 1, 9) = Format(SubCustom(4), "##,###")
                    .TextMatrix(.Rows - 1, 10) = Format(SubCustom(5), "##,###")
                    
                    
                    .Rows = .Rows + 1
                    .RowHeight(.Rows - 1) = 350
                    .MergeCells = flexMergeFree
                    .MergeRow(.Rows - 1) = True
                    .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
                    .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 3) = "공급가액:  " & Format(SubCustom(4), "##,##0") & "원"
                    .Cell(flexcpText, .Rows - 1, 4, .Rows - 1, 6) = "부가세:  " & Format(SubCustom(5), "##,##0") & "원"
                    .Cell(flexcpText, .Rows - 1, 7, .Rows - 1, 10) = "총금액:  " & Format(SubCustom(4) + SubCustom(5), "##,##0") & "원"
                    
                    For iCnt = 1 To 5
                        SubArticle(iCnt) = 0
                        SubCustom(iCnt) = 0
                    Next iCnt
                    
                    sCustom = rs!kCustom
                    sArticle = rs!Article
                Else
                    If sArticle <> rs!Article Then
                        .Rows = .Rows + 1
                        .RowHeight(.Rows - 1) = 300

                        .TextMatrix(.Rows - 1, 0) = ""
                        .TextMatrix(.Rows - 1, 1) = sCustom
                        .TextMatrix(.Rows - 1, 2) = sArticle
                        .Cell(flexcpText, .Rows - 1, 3) = "품명 계"
                        .TextMatrix(.Rows - 1, 5) = Format(SubArticle(1), "##,###")
                        .TextMatrix(.Rows - 1, 7) = Format(SubArticle(2), "##,###")
                        .TextMatrix(.Rows - 1, 8) = Format(SubArticle(3), "##,###")
                        .TextMatrix(.Rows - 1, 9) = Format(SubArticle(4), "##,###")
                        .TextMatrix(.Rows - 1, 10) = Format(SubArticle(5), "##,###")
                        
                        For iCnt = 1 To 5
                            SubArticle(iCnt) = 0
                        Next iCnt
                        
                        sArticle = rs!Article
                    End If
                End If
                
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = 300
    
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = rs!kCustom
                .TextMatrix(.Rows - 1, 2) = rs!Article
                .TextMatrix(.Rows - 1, 3) = rs!WorkName
                .TextMatrix(.Rows - 1, 4) = rs!OrderNo
                .TextMatrix(.Rows - 1, 5) = Format(rs!OrderQty, "##,###")
                .TextMatrix(.Rows - 1, 6) = rs!UnitPrice
                .TextMatrix(.Rows - 1, 7) = Format(rs!outjwqty, "##,###")
                .TextMatrix(.Rows - 1, 8) = Format(rs!OutQty, "##,###")
                .TextMatrix(.Rows - 1, 9) = Format(rs!UnitPrice * rs!OutQty, "##,###")
                .TextMatrix(.Rows - 1, 10) = Format(rs!UnitPrice * rs!OutQty * 0.1, "##,###")
                
                SubArticle(1) = SubArticle(1) + rs!OrderQty
                SubCustom(1) = SubCustom(1) + rs!OrderQty
                
                SubArticle(2) = SubArticle(2) + rs!outjwqty
                SubCustom(2) = SubCustom(2) + rs!outjwqty
                
                SubArticle(3) = SubArticle(3) + rs!OutQty
                SubCustom(3) = SubCustom(3) + rs!OutQty
                
                SubArticle(4) = SubArticle(4) + rs!OutQty * rs!UnitPrice
                SubCustom(4) = SubCustom(4) + rs!OutQty * rs!UnitPrice
                
                SubArticle(5) = SubArticle(5) + rs!OutQty * rs!UnitPrice * 0.1
                SubCustom(5) = SubCustom(5) + rs!OutQty * rs!UnitPrice * 0.1
                
                .TextMatrix(.Rows - 1, 21) = rs!CustomID
                .TextMatrix(.Rows - 1, 22) = rs!ArticleID
                .TextMatrix(.Rows - 1, 23) = rs!WorkID
                
                rs.MoveNext
            Next i
            
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            .TextMatrix(.Rows - 1, 0) = ""
            .TextMatrix(.Rows - 1, 1) = sCustom
            .TextMatrix(.Rows - 1, 2) = sArticle
            .Cell(flexcpText, .Rows - 1, 3) = "품명 계"
            .TextMatrix(.Rows - 1, 5) = Format(SubArticle(1), "##,###")
            .TextMatrix(.Rows - 1, 7) = Format(SubArticle(2), "##,###")
            .TextMatrix(.Rows - 1, 8) = Format(SubArticle(3), "##,###")
            .TextMatrix(.Rows - 1, 9) = Format(SubArticle(4), "##,###")
            .TextMatrix(.Rows - 1, 10) = Format(SubArticle(5), "##,###")
        
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            .TextMatrix(.Rows - 1, 0) = ""
            .TextMatrix(.Rows - 1, 1) = sCustom
            .Cell(flexcpText, .Rows - 1, 2, .Rows - 1, 2) = "거래처 계"
            .TextMatrix(.Rows - 1, 5) = Format(SubCustom(1), "##,###")
            .TextMatrix(.Rows - 1, 7) = Format(SubCustom(2), "##,###")
            .TextMatrix(.Rows - 1, 8) = Format(SubCustom(3), "##,###")
            .TextMatrix(.Rows - 1, 9) = Format(SubCustom(4), "##,###")
            .TextMatrix(.Rows - 1, 10) = Format(SubCustom(5), "##,###")
            
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 350
            .MergeCells = flexMergeFree
            .MergeRow(.Rows - 1) = True
            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            .Cell(flexcpText, .Rows - 1, 1, .Rows - 1, 3) = "공급가액:  " & Format(SubCustom(4), "##,##0") & "원"
            .Cell(flexcpText, .Rows - 1, 4, .Rows - 1, 6) = "부가세:  " & Format(SubCustom(5), "##,##0") & "원"
            .Cell(flexcpText, .Rows - 1, 7, .Rows - 1, 10) = "총금액:  " & Format(SubCustom(4) + SubCustom(5), "##,##0") & "원"
            
            
        End If
        rs.Close
        Set rs = Nothing
        
        .MergeCells = flexMergeFree
        .MergeCol(1) = True
        .MergeCol(2) = True
        
        .Redraw = flexRDDirect
        .Row = 0
    End With
    
                
    
    Exit Sub

ErrHandler:
    Set oSubul = Nothing
    Set rs = Nothing
    Call ErrorBox(Err.Number, "frmBillView.FillGridBill", Err.Description)
End Sub


