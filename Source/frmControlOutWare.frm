VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmControlOutWare 
   Caption         =   "출고 조정"
   ClientHeight    =   9225
   ClientLeft      =   3090
   ClientTop       =   3060
   ClientWidth     =   11850
   Icon            =   "frmControlOutWare.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   11850
   Begin VSFlex7LCtl.VSFlexGrid grdOrder 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   11850
      _cx             =   20902
      _cy             =   14208
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
   Begin Threed.SSPanel pnlSearch 
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   1270
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Frame fraOrder 
         Height          =   435
         Left            =   6300
         TabIndex        =   26
         Top             =   -60
         Width           =   2625
         Begin VB.OptionButton optOrder 
            Caption         =   "OrderNo"
            Height          =   195
            Index           =   1
            Left            =   1410
            TabIndex        =   28
            Top             =   180
            Width           =   1035
         End
         Begin VB.OptionButton optOrder 
            Caption         =   "관리번호"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   180
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin VB.Frame fraDate 
         Height          =   795
         Left            =   30
         TabIndex        =   18
         Top             =   -90
         Width           =   1095
         Begin VB.OptionButton optDate 
            Caption         =   "일자무관"
            Height          =   180
            Index           =   0
            Left            =   30
            TabIndex        =   21
            Top             =   150
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optDate 
            Caption         =   "수주일자"
            Height          =   180
            Index           =   1
            Left            =   30
            TabIndex        =   20
            Top             =   360
            Width           =   1035
         End
         Begin VB.OptionButton optDate 
            Caption         =   "납기일자"
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   19
            Top             =   570
            Width           =   1035
         End
      End
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   4380
         TabIndex        =   6
         Top             =   45
         Width           =   1545
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4380
         TabIndex        =   5
         Top             =   390
         Width           =   1545
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금월"
         Height          =   315
         Index           =   1
         Left            =   1140
         MousePointer    =   99  '사용자 정의
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   540
      End
      Begin VB.CommandButton cmdTerm 
         Caption         =   "금일"
         Height          =   315
         Index           =   0
         Left            =   1140
         MousePointer    =   99  '사용자 정의
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   540
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   7440
         TabIndex        =   2
         Top             =   390
         Width           =   1485
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   30
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23724033
         CurrentDate     =   36871
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   1680
         TabIndex        =   8
         Top             =   390
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23724033
         CurrentDate     =   36871
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   1
         Left            =   3390
         TabIndex        =   9
         Top             =   45
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   529
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "거래처"
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   45
            Width           =   885
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   0
         Left            =   5940
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   196609
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   3
         Left            =   3390
         TabIndex        =   12
         Top             =   390
         Width           =   990
         _ExtentX        =   1746
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
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "품   명"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   885
         End
      End
      Begin Threed.SSCommand cmdFind 
         Height          =   300
         Index           =   1
         Left            =   5940
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   390
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
      Begin Threed.SSCommand cmdSearch 
         Height          =   660
         Left            =   8940
         TabIndex        =   15
         Top             =   30
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1164
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
         Caption         =   "        검색"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   660
         Left            =   10860
         TabIndex        =   16
         Tag             =   "PERM_ADDNEW"
         Top             =   30
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1164
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
      Begin Threed.SSCommand cmdSave 
         Height          =   660
         Left            =   9900
         TabIndex        =   17
         Top             =   30
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1164
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
         Caption         =   "        적용"
         PictureAlignment=   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   0
         Left            =   6300
         TabIndex        =   22
         Top             =   405
         Width           =   1140
         _ExtentX        =   2011
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
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkSearch 
            Caption         =   "관리번호"
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1065
         End
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   180
         Index           =   1
         Left            =   2955
         TabIndex        =   25
         Top             =   465
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   180
         Index           =   0
         Left            =   2955
         TabIndex        =   24
         Top             =   105
         Width           =   360
      End
   End
   Begin Threed.SSCommand cmdCheck 
      Height          =   480
      Index           =   0
      Left            =   0
      TabIndex        =   29
      ToolTipText     =   "모두 체크처리합니다"
      Top             =   720
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   847
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
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand cmdCheck 
      Height          =   480
      Index           =   1
      Left            =   570
      TabIndex        =   30
      ToolTipText     =   "모두 체크를 해제합니다"
      Top             =   720
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   847
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
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   435
      Index           =   1
      Left            =   7770
      TabIndex        =   31
      Top             =   750
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   767
      _Version        =   196609
      Caption         =   "※  체크처리 된것만 변경내용이 저장됩니다."
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlCount 
      Height          =   435
      Left            =   4410
      TabIndex        =   32
      Top             =   750
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   767
      _Version        =   196609
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   435
      Index           =   0
      Left            =   1230
      TabIndex        =   33
      Top             =   750
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   767
      _Version        =   196609
      Font3D          =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "■  마감이 안된 수주건 조회"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "frmControlOutWare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click(Index As Integer)
    Dim SetValue, i%
    
    If Index = 0 Then   '[0] 전체선택
        SetValue = flexChecked
    Else                '[1] 선택 해제
        SetValue = flexUnchecked
    End If

    With grdOrder
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, 0) > 0 Then
                .Cell(flexcpChecked, i, 0) = SetValue
            End If
        Next i
    End With

End Sub

Private Sub cmdSave_Click()
    If CheckData() Then
        If MsgBox("선택한 건에 대해서 출고 절수와 수량을 적용시키겠습니까?" & vbCrLf & vbCrLf & _
                  "출고일자는 오늘날짜(" & Format(Now, "YYYY/MM/DD") & ")로 적용됩니다", vbQuestion + vbYesNo, "적용 여부") = vbYes Then
            If SaveData() Then
                MsgBox "출고 절수와 수량을 적용하였습니다.", vbInformation + vbOKOnly, "출고처리 완료"
                Call FillGrid
            End If
        
        End If
    End If
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    If Index = 0 Then '[1] 금일
        dtpDate(0) = Date
        dtpDate(1) = Date
    ElseIf Index = 1 Then '[2] 금월
        dtpDate(0) = DateSerial(Year(Date), Month(Date), 1)
        dtpDate(1) = Date
    End If
End Sub

Private Sub Form_Activate()
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub Form_Load()
    Dim i%
    
    Me.Move 0, 0, 11975, 9660
    
    Call SetOperate(Me)
    For i = 0 To 1
        cmdFind(i).Picture = LoadResPicture("FIND", vbResIcon)
        cmdFind(i).Enabled = False
        dtpDate(i) = Now
    Next i
    cmdSave.Picture = LoadResPicture("COMMAND", vbResIcon)
    cmdCheck(0).Picture = LoadResPicture("SAVE", vbResIcon)
    cmdCheck(1).Picture = LoadResPicture("CANCEL", vbResIcon)
    Call InitGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlusMDI.pnlMenu.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub chkSearch_Click(Index As Integer)
    If chkSearch(Index).Value = vbChecked Then
        txtSearch(Index).Enabled = True
        txtSearch(Index).SetFocus
        If Index < 2 Then
            cmdFind(Index).Enabled = True
        End If
    Else
        txtSearch(Index).Enabled = False
        If Index < 2 Then
            cmdFind(Index).Enabled = False
        End If
    End If
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index = 0 Then
        Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(0))
    ElseIf Index = 1 Then
        Call ReturnCode(LG_ARTICLE, , False, txtSearch(1))
    End If

End Sub

Private Sub cmdSearch_Click()
    Call FillGrid
End Sub

Private Function SaveData() As Boolean
    Dim oSubul As PlusLib2.CSubul
    Dim tWork() As PlusLib2.TOutWareRec
    Dim tWorkSub() As PlusLib2.TOutWareSubRec
    Dim iRow%, iCntChk%, iCount%
    Dim dRate As Double
    Dim sDate$, sTime$
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    SaveData = False
    
    sDate = Format(Date, "YYYYMMDD")
    sTime = Format(time, "HHMM")
    
    With grdOrder
        For iRow = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, iRow, 0) = flexChecked Then
                iCntChk = iCntChk + 1
            End If
        Next iRow
        
        ReDim tWork(iCntChk)
        ReDim tWorkSub(iCntChk)
        
    
        Set oSubul = New PlusLib2.CSubul
        oSubul.Connection = g_adoCon
        oSubul.UserName = g_sUserName
        
        iCount = 0
        For iRow = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, iRow, 0) = flexChecked Then
                dRate = CDbl("0" & .TextMatrix(iRow, 26)) + CDbl("0" & .TextMatrix(iRow, 27))
            
                tWork(iCount).OrderID = .TextMatrix(iRow, 24)
                tWork(iCount).OutSeq = 0
                tWork(iCount).OutClss = "1"     ' 정상출고
                tWork(iCount).WorkID = .TextMatrix(iRow, 23)
                tWork(iCount).ExchRate = 0
                tWork(iCount).UnitPrice = 0
                tWork(iCount).OutCustom = ""
                tWork(iCount).LossRate = 0
                tWork(iCount).LossQty = 0
                tWork(iCount).OutRoll = CInt("0" & .TextMatrix(iRow, 11))
                tWork(iCount).OutQty = CLng(.TextMatrix(iRow, 12))
                tWork(iCount).OutRealQty = CLng(tWork(iCount).OutQty * (1 + (dRate / 100)))
                tWork(iCount).OutDate = sDate
                tWork(iCount).ResultDate = sDate
                tWork(iCount).OutTime = sTime
                tWork(iCount).BoOutClss = ""
                tWork(iCount).BoConfirmClss = ""
                tWork(iCount).BoConfirmDate = ""
                tWork(iCount).LoadTime = sTime
                tWork(iCount).TranNo = ""
                tWork(iCount).TranSeq = 0
                tWork(iCount).TelNo = ""
                tWork(iCount).Remark = "정산처리분"
                tWork(iCount).OutType = "1"
                
                tWorkSub(iCount).OrderID = .TextMatrix(iRow, 24)
                tWorkSub(iCount).OutSeq = 0
                tWorkSub(iCount).OutSubSeq = 1
                tWorkSub(iCount).OrderSeq = CInt(.TextMatrix(iRow, 30))
                tWorkSub(iCount).RollSeq = 1
                tWorkSub(iCount).LotNo = ""
                tWorkSub(iCount).OutQty = CLng(.TextMatrix(iRow, 12))

                
                iCount = iCount + 1
            End If
        Next iRow
    End With
    
    If Not oSubul.AddNewOutData(tWork(), tWorkSub()) Then
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
    Call ErrorBox(Err.Number, "frmControlOutWare.SaveData", Err.Description)
End Function

Private Function CheckData() As Boolean
Dim iRow%, iChkCnt%

    CheckData = True
    
    With grdOrder
        If .Rows = .FixedRows Then
            MsgBox "해당 항목을 체크해야 적용됩니다", vbInformation + vbOKOnly, "항목 선택"
            CheckData = False
            Exit Function
        End If
        For iRow = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, iRow, 0) = flexChecked Then
                iChkCnt = iChkCnt + 1
                If Not (IsNumeric(.TextMatrix(iRow, 11))) Or _
                    Not (IsNumeric(.TextMatrix(iRow, 12))) Then
                    MsgBox "숫자가 아닌 데이터가 포함되어 있습니다" & vbCrLf & vbCrLf & _
                            "확인후 다시 작업하여 주십시요", vbCritical + vbOKOnly, "숫자 입력"
                    CheckData = False
                    Exit Function
                End If
            End If
        Next iRow
        If iChkCnt = 0 Then
            MsgBox "적어도 한 건은 체크해야 적용됩니다", vbInformation + vbOKOnly, "항목 선택"
            CheckData = False
            Exit Function
        End If
        
    End With

End Function

Private Sub InitGrid()
    Dim iCol%
    
    With grdOrder
        .Redraw = flexRDNone
        
        .SelectionMode = flexSelectionFree
        .HighLight = flexHighlightNever
        .ScrollBars = flexScrollBarVertical
        .FocusRect = flexFocusSolid
        .ScrollTrack = True
        .WordWrap = False
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
        
        .Cols = 31:     .Rows = 3
        .FixedCols = 0: .FixedRows = 3
        
        .RowHeight(0) = 0
        .RowHeight(1) = 0
        .RowHeight(2) = 300

        For iCol = 0 To .Cols - 1
            .ColWidth(iCol) = 0
            .FixedAlignment(iCol) = flexAlignCenterCenter
        Next iCol
        
        .TextMatrix(2, 0) = "":                 .ColWidth(0) = 300:         .ColAlignment(0) = flexAlignCenterCenter
        .TextMatrix(2, 1) = "거래처":           .ColWidth(1) = 1300:        .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(2, 2) = "품명":             .ColWidth(2) = 2000:        .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(2, 3) = "색상":             .ColWidth(3) = 1800:        .ColAlignment(3) = flexAlignLeftCenter
        .TextMatrix(2, 4) = "가공구분":         .ColWidth(4) = 1100:        .ColAlignment(4) = flexAlignLeftCenter
        .TextMatrix(2, 5) = "관리번호":         .ColWidth(5) = 1300:        .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(2, 6) = "OrderNo":          .ColWidth(6) = 0:           .ColAlignment(6) = flexAlignLeftCenter
        .TextMatrix(2, 7) = "수주량":           .ColWidth(7) = 900:         .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(2, 8) = "수주량":           .ColWidth(8) = 500:         .ColAlignment(8) = flexAlignLeftCenter
        .TextMatrix(2, 9) = "축율+Loss":        .ColWidth(9) = 1000:        .ColAlignment(9) = flexAlignCenterCenter
        .TextMatrix(2, 10) = "입고량":          .ColWidth(10) = 0:          .ColAlignment(10) = flexAlignRightCenter
        .TextMatrix(2, 11) = "출고절수":        .ColWidth(11) = 0:          .ColAlignment(11) = flexAlignRightCenter
        .TextMatrix(2, 12) = "출고수량":        .ColWidth(12) = 1000:       .ColAlignment(12) = flexAlignRightCenter
        .TextMatrix(2, 13) = "수주일자":        .ColWidth(13) = 0:          .ColAlignment(13) = flexAlignCenterCenter
        .TextMatrix(2, 14) = "납기일자":        .ColWidth(14) = 0:          .ColAlignment(14) = flexAlignCenterCenter
        .TextMatrix(2, 15) = "비고":            .ColWidth(15) = 0:          .ColAlignment(15) = flexAlignLeftCenter
        
        .TextMatrix(2, 21) = "CustomID"
        .TextMatrix(2, 22) = "ArticleID"
        .TextMatrix(2, 23) = "WorkID"
        .TextMatrix(2, 24) = "OrderID"
        .TextMatrix(2, 25) = "UnitClss"
        .TextMatrix(2, 26) = "ChunkRate"
        .TextMatrix(2, 27) = "LossRate"
        .TextMatrix(2, 28) = "AcptDate"
        .TextMatrix(2, 29) = "DvlyDate"
        .TextMatrix(2, 30) = "OrderSeq"
        
        .MergeCells = flexMergeFixedOnly
        .MergeRow(2) = True
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FillGrid()
    Dim oSubul As PlusLib2.CSubul
    Dim rs As Recordset
    Dim i%
    
    On Error GoTo ErrHandler
    
    pnlCount.Caption = ""
    Set oSubul = New PlusLib2.CSubul
    oSubul.Connection = g_adoCon
    
    Set rs = oSubul.GetNotCloseOrder(IIf(optDate(0).Value = True, 0, IIf(optDate(1).Value = True, 1, 2)), _
                                    MakeDate(DF_SHORT, dtpDate(0)), MakeDate(DF_SHORT, dtpDate(1)), _
                                    IIf(chkSearch(0) = vbChecked, 1, 0), txtSearch(0).Tag, _
                                    IIf(chkSearch(1) = vbChecked, 1, 0), txtSearch(1).Tag, _
                                    IIf(chkSearch(2) = vbChecked, IIf(optOrder(0).Value = True, 1, 2), 0), Trim(txtSearch(2)))
    
    Set oSubul = Nothing
        
    pnlCount.Caption = "총 검색 건수 : " & CStr(rs.RecordCount) & "건"
    With grdOrder
        .Redraw = flexRDNone
        .Rows = .FixedRows
    
        For i = 1 To rs.RecordCount
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = 300
            
            .Cell(flexcpChecked, .Rows - 1, 0) = flexUnchecked
            .TextMatrix(.Rows - 1, 1) = Trim(rs!kCustom)
            .TextMatrix(.Rows - 1, 2) = Trim(rs!Article)
            .TextMatrix(.Rows - 1, 3) = Trim(rs!Color)
            .TextMatrix(.Rows - 1, 4) = Trim(rs!WorkName)
            .TextMatrix(.Rows - 1, 5) = MakeOrderID(rs!OrderID, OM_EXPAND)
            .TextMatrix(.Rows - 1, 6) = Trim(rs!OrderNo)
            .TextMatrix(.Rows - 1, 7) = Format(rs!ColorQty, "##,##0")
            .TextMatrix(.Rows - 1, 8) = IIf(rs!UnitClss = "0", "YDS", "MTS")
            .TextMatrix(.Rows - 1, 9) = CStr(rs!ChunkRate) & "+" & CStr(rs!LossRate)
'            .TextMatrix(.Rows - 1, 10) = Format(rs!InQty, "##,##0")
            .TextMatrix(.Rows - 1, 11) = Format(rs!OutRoll, "##,##0")
            .TextMatrix(.Rows - 1, 12) = Format(rs!OutQty, "##,##0")
            .TextMatrix(.Rows - 1, 13) = rs!AcptDate
            .TextMatrix(.Rows - 1, 14) = rs!DvlyDate
            .TextMatrix(.Rows - 1, 15) = rs!Remark
            
            .TextMatrix(.Rows - 1, 21) = rs!CustomID
            .TextMatrix(.Rows - 1, 22) = rs!ArticleID
            .TextMatrix(.Rows - 1, 23) = rs!WorkID
            .TextMatrix(.Rows - 1, 24) = rs!OrderID
            .TextMatrix(.Rows - 1, 25) = rs!UnitClss
            .TextMatrix(.Rows - 1, 26) = rs!ChunkRate
            .TextMatrix(.Rows - 1, 27) = rs!LossRate
            .TextMatrix(.Rows - 1, 28) = rs!AcptDate
            .TextMatrix(.Rows - 1, 29) = rs!DvlyDate
            .TextMatrix(.Rows - 1, 30) = rs!OrderSeq
            
            rs.MoveNext
        Next i
        rs.Close
        
        If .Rows > .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, 12, .Rows - 1, 12) = &HC0FFC0
        End If
        Set rs = Nothing
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
        End If
        
        .Redraw = flexRDDirect
    End With
    
    Exit Sub

ErrHandler:
    Set rs = Nothing
    Set oSubul = Nothing
    Call ErrorBox(Err.Number, "frmControlOutWare.FillGrid", Err.Description)
End Sub


Private Sub grdOrder_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grdOrder
'        If Col = 11 Or Col = 12 Then
        If Col = 12 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) Then
                MsgBox "숫자를 정확히 입력하여 주십시요", vbExclamation + vbOKOnly, "숫자 입력"
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(CLng(.TextMatrix(Row, Col)), "##,##0")
        End If
    End With
End Sub

Private Sub grdOrder_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With grdOrder
'        If Col = 0 Or Col = 11 Or Col = 12 Then
        If Col = 0 Or Col = 12 Then
            Cancel = False
        Else
            Cancel = True
        End If
'        If Col = 11 Or Col = 12 Then
        If Col = 12 Then
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusNone
        End If
    End With
End Sub

Private Sub grdOrder_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With grdOrder
'        If Col >= 11 And Col <= 12 And KeyAscii = vbKeyReturn Then
        If Col = 12 And KeyAscii = vbKeyReturn Then
'            If Not IsNumeric(.TextMatrix(Row, Col)) Then
'                MsgBox "숫자를 정확히 입력하여 주십시요", vbExclamation + vbOKOnly, "숫자 입력"
'                .Row = Row
'                .Col = Col
'                Exit Sub
'            End If
'            .TextMatrix(Row, Col) = Format(CLng(.TextMatrix(Row, Col)), "##,##0")
        
            Select Case Col
'                Case 11:
'                    .Col = 12
                Case 12:
                    If Row < .Rows - 1 Then
                        .Row = .Row + 1
'                        .Col = 11
                        .Col = 12
                    End If
            End Select
        End If
    End With
End Sub

Private Sub optOrder_Click(Index As Integer)
    If optOrder(Index).Value = True Then
        chkSearch(2).Caption = optOrder(Index).Caption
        If Index = 0 Then
            grdOrder.ColWidth(5) = 1300
            grdOrder.ColWidth(6) = 0
        Else
            grdOrder.ColWidth(5) = 0
            grdOrder.ColWidth(6) = 1300
        End If
    End If
End Sub

Private Sub txtSearch_KeyPress(Index As Integer, KeyAscii As Integer)
    If chkSearch(Index).Value = 1 Then
        If KeyAscii = vbKeyReturn Then
            If Index = 0 Then
                Call ReturnCode(LG_CUSTOM, 0, False, txtSearch(Index))
            ElseIf Index = 1 Then
                Call ReturnRef(LG_ARTICLE, , False, txtSearch(Index))
            End If
            
        End If
    End If
End Sub
