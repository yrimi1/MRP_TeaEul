VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmArticle 
   Caption         =   "ǰ��ã��"
   ClientHeight    =   5745
   ClientLeft      =   5310
   ClientTop       =   4335
   ClientWidth     =   7635
   Icon            =   "frmArticle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7635
   Begin Threed.SSCommand cmdExit 
      Height          =   660
      Left            =   5910
      TabIndex        =   19
      Top             =   5040
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1164
      _Version        =   196609
      Caption         =   "      ���(&X)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel pnlBoard 
      Height          =   4920
      Left            =   3420
      TabIndex        =   6
      Top             =   45
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   8678
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel pnlEdit 
         Height          =   810
         Left            =   45
         TabIndex        =   7
         Top             =   915
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   1429
         _Version        =   196609
         Enabled         =   0   'False
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboDye 
            Height          =   300
            Left            =   1140
            TabIndex        =   29
            Top             =   1875
            Width           =   2835
         End
         Begin PlusFind2.WizText txtCode 
            Height          =   300
            Left            =   1140
            TabIndex        =   14
            Top             =   75
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   4
            BackColor       =   12648384
         End
         Begin PlusFind2.WizText txtName 
            Height          =   300
            Index           =   0
            Left            =   1140
            TabIndex        =   15
            Top             =   435
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   35
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   0
            Left            =   75
            TabIndex        =   8
            Top             =   75
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��   ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   1
            Left            =   75
            TabIndex        =   9
            Top             =   435
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "ǰ   ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   2
            Left            =   75
            TabIndex        =   10
            Top             =   825
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��   ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PlusFind2.WizText txtName 
            Height          =   300
            Index           =   1
            Left            =   1140
            TabIndex        =   16
            Top             =   825
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   20
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   3
            Left            =   75
            TabIndex        =   22
            Top             =   1155
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "������"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PlusFind2.WizText txtName 
            Height          =   300
            Index           =   2
            Left            =   1140
            TabIndex        =   23
            Top             =   1155
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   20
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   4
            Left            =   75
            TabIndex        =   24
            Top             =   1875
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "������"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel pnlCaption 
            Height          =   300
            Index           =   5
            Left            =   75
            TabIndex        =   25
            Top             =   1515
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   196609
            Caption         =   "��   ��"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PlusFind2.WizText txtName 
            Height          =   300
            Index           =   3
            Left            =   1140
            TabIndex        =   26
            Top             =   1515
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   20
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   0
            Left            =   3675
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   825
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            PictureFrames   =   1
            Enabled         =   0   'False
            Picture         =   "frmArticle.frx":000C
            ButtonStyle     =   3
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdFind 
            Height          =   300
            Index           =   1
            Left            =   3675
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1140
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   196609
            PictureFrames   =   1
            Enabled         =   0   'False
            Picture         =   "frmArticle.frx":0326
            ButtonStyle     =   3
            Outline         =   0   'False
         End
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "���(&C)"
         Height          =   780
         Index           =   4
         Left            =   900
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   18
         ToolTipText     =   "�ڷ� ���"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&U)"
         Height          =   780
         Index           =   1
         Left            =   2490
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   12
         ToolTipText     =   "�ڷ� ����"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&D)"
         Height          =   780
         Index           =   2
         Left            =   3285
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   13
         ToolTipText     =   "�ڷ� ����"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "�߰�(&A)"
         Height          =   780
         Index           =   0
         Left            =   1695
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   11
         ToolTipText     =   "�ڷ� �߰�"
         Top             =   60
         Width           =   780
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&S)"
         Height          =   780
         Index           =   3
         Left            =   120
         MousePointer    =   99  '����� ����
         Style           =   1  '�׷���
         TabIndex        =   17
         ToolTipText     =   "�ڷ� ����"
         Top             =   60
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSPanel pnlMsg 
         Height          =   510
         Left            =   240
         TabIndex        =   21
         Top             =   4110
         Visible         =   0   'False
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   900
         _Version        =   196609
         BackColor       =   65535
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlSearch 
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   1588
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton optSize 
         Caption         =   "��"
         Height          =   330
         Index           =   1
         Left            =   2655
         Style           =   1  '�׷���
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton optSize 
         Caption         =   "���"
         Height          =   330
         Index           =   0
         Left            =   2655
         Style           =   1  '�׷���
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   645
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   75
         TabIndex        =   2
         Top             =   465
         Width           =   2025
      End
      Begin Threed.SSPanel pnlCaption 
         Height          =   300
         Index           =   25
         Left            =   90
         TabIndex        =   1
         Top             =   120
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         _Version        =   196609
         Caption         =   "ǰ�� �˻�"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdAll 
         Height          =   330
         Left            =   2130
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   450
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   582
         _Version        =   196609
         MousePointer    =   99
         CaptionStyle    =   1
         PictureAnimationEnabled=   0   'False
         Alignment       =   6
         PictureAlignment=   0
         BevelWidth      =   1
         ShapeSize       =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdData 
      Height          =   3990
      Left            =   15
      TabIndex        =   31
      Top             =   975
      Width           =   3360
      _cx             =   5927
      _cy             =   7038
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Height          =   660
      Left            =   4140
      TabIndex        =   30
      Top             =   5040
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1164
      _Version        =   196609
      Caption         =   "      ����(&Q)"
      PictureAlignment=   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "�˻��Ǽ� :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   105
      TabIndex        =   20
      Top             =   5190
      Width           =   945
   End
End
Attribute VB_Name = "frmArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIMIT_ROW = 14
Private Const LIMIT_WIDTH = 2400

Dim m_bSelected     As Boolean
Dim wData()
'------------------------------------------------------------------
Dim m_sFlag        As String * 1
Dim m_bSortForward As Boolean
Dim m_bSkip As Boolean

Public Function SetMsg(SelData(), Optional sNewData) As Boolean
    Dim i%
      
    If IsMissing(sNewData) Then
        Me.Show vbModal
    Else
        If sNewData = "" Then
           
            Me.Show vbModal
        Else
            Call SetGrid(FL_BY_CODE, sNewData)
            If grdData.Rows = grdData.FixedRows Then
                txtSearch = sNewData
                Call SetGrid(FL_BY_NAME, sNewData)
            End If
            
            '------------------------------------------------
            With grdData
                If .Rows > .FixedRows Then
                    If .Rows = .FixedRows + 1 Then
                        Call SelectData
                    Else
                        Me.Show vbModal
                    End If
                Else
                    If MsgBox(LoadResString(112), vbQuestion + vbYesNo) = vbYes Then
                        Call cmdOperate_Click(ID_ADDNEW)
                        txtName(0).Text = sNewData
                        
                        Me.Show vbModal
                    End If
                End If
            End With
        End If
    End If
    
    '=====================================================================
    If m_bSelected Then
        With grdData
            ReDim SelData(UBound(wData))
            For i = LBound(wData) To UBound(wData)
                SelData(i) = wData(i)
            Next i
        End With
    End If
    
    SetMsg = m_bSelected
End Function

Private Sub SetGrid(ByVal Index As EFindClss, Optional sNewData)
    Dim oArticle As PlusLib2.CArticle
    Dim rs As ADODB.Recordset
    
    Dim nNowRow&, sID$
    
    On Error GoTo ErrHandler
    
    m_bSkip = True
       
    Set oArticle = New PlusLib2.CArticle
    oArticle.Connection = adoCon
    
    If Index = FL_BY_CODE Then
        If LenB(StrConv(sNewData, vbFromUnicode)) < 6 Then
            Set rs = oArticle.GetArticle(CStr(sNewData))
        Else
            Set oArticle = Nothing
            Exit Sub
        End If
    ElseIf Index = FL_BY_NAME Then
        Set rs = oArticle.GetArticle(sNewData)
    End If
    Set oArticle = Nothing
    
    With grdData
        .Redraw = flexRDNone
        
        If .Rows > .FixedRows Then
            If m_sFlag = ID_ADDNEW Then
                nNowRow = .Rows
            Else
                nNowRow = .Row
            End If
            .Rows = .FixedRows
        Else
            nNowRow = 1
        End If
        
        Do Until rs.EOF
            .AddItem CStr(.Rows) & vbTab & CStr(rs!ArticleID) & vbTab & rs!Article & vbTab & _
                CheckNull(rs!Thread) & vbTab & CheckNull(rs!ThreadID) & vbTab & _
                CheckNull(rs!StuffWidth) & vbTab & CheckNull(rs!StuffWidthID) & vbTab & _
                CheckNull(rs!DyeingID) & vbTab & CheckNull(rs!Weight)
            
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        lblCount.Caption = "�˻��Ǽ� : " & CStr(.Rows - 1) & " ��"
        
        If .Rows > .FixedRows Then
            If .Rows > nNowRow Then
                .Row = nNowRow
            Else
                .Row = .Rows - 1
            End If
            .TopRow = .Row
            .Col = .FixedCols
            .ColSel = .Cols - 1
            
            .HighLight = flexHighlightAlways
            
            Call ShowData
            
        Else
            .HighLight = flexHighlightNever
            
            Call ClearData
        End If
        
        .Redraw = flexRDDirect
    End With
    
    m_bSkip = False
    Exit Sub
ErrHandler:
    Set oArticle = Nothing
    Set rs = Nothing
    
    Call ErrorBox(Err.Number, "Article.SetGrid", Err.Description)
End Sub

Private Sub SelectData()
    Dim i%
    
    On Error Resume Next
    
    If grdData.Rows > 1 Then
        m_bSelected = True
        
        ReDim wData(grdData.Cols - 2)
        With grdData
            For i = 1 To .Cols - 1
                wData(i - 1) = .TextMatrix(.Row, i)
            Next i
        End With
        
        Me.Hide
    End If
End Sub


Private Sub cmdAll_Click()
    Dim iLoop As Integer

    With grdData
        .Redraw = flexRDNone

        For iLoop = .FixedRows To .Rows - .FixedRows
            .RowHidden(iLoop) = False
        Next iLoop

        .Redraw = flexRDDirect
    End With

    txtSearch.Text = ""
    cmdAll.Visible = False
End Sub

Private Sub cmdExit_Click()
    m_bSelected = False
    
    Me.Hide
End Sub
'
'Private Sub cmdFind_Click(Index As Integer)
'    If Index = 0 Then
'        Call ReturnCode(LG_THREAD, , False, txtName(1))
'    Else
'        Call ReturnCode(LG_STUFFWIDTH, , False, txtName(2))
'    End If
'End Sub

Private Sub cmdOperate_Click(Index As Integer)

    On Error GoTo ErrHandler
    
    Select Case Index
        Case ID_ADDNEW
            m_sFlag = ID_ADDNEW
            Call ClearData
            Call ChangeMode(Me, False)
            
            pnlEdit.Enabled = True
            txtCode.Locked = False
'            txtName(0).SetFocus
            pnlMsg.Caption = LoadResString(302)
        Case ID_UPDATE
            If grdData.Rows > grdData.FixedRows Then
                m_sFlag = ID_UPDATE
                Call ChangeMode(Me, False)
                pnlEdit.Enabled = True
                
                txtCode.Locked = True
                
                cmdFind(0).Enabled = True
                cmdFind(1).Enabled = True
                
                txtName(0).SetFocus
                pnlMsg.Caption = LoadResString(303)
            End If
        Case ID_DELETE
            If MsgBox(LoadResString(201), vbQuestion + vbYesNo, "����Ȯ��") = vbYes Then
                m_sFlag = ID_DELETE
                If SaveData Then
                    Call SetGrid(FL_BY_NAME)
                    If Len(txtSearch) > 0 Then Call txtSearch_Change
                    m_sFlag = ""
                End If
            End If
        Case ID_SAVE
            If Not CheckData() Then Exit Sub
    
            If SaveData() Then
                Call ChangeMode(Me, True)
                pnlEdit.Enabled = False
                cmdFind(0).Enabled = False
                cmdFind(1).Enabled = False
                Call SetGrid(FL_BY_NAME)
                If Len(txtSearch) > 0 Then Call txtSearch_Change
                m_sFlag = ""
            End If
            grdData.SetFocus
        Case ID_CANCEL
            m_sFlag = ""
            Call ChangeMode(Me, True)
            pnlEdit.Enabled = False
            
            cmdFind(0).Enabled = False
            cmdFind(1).Enabled = False
            With grdData
                If .Rows > .FixedRows Then
                    Call ShowData
                Else
                    Call ClearData
                End If
            End With
            grdData.SetFocus
            
        End Select

    Exit Sub
ErrHandler:
    MsgBox Err.Number & ":" & Err.Description, vbCritical
    Err.Clear
End Sub

Private Function CheckData() As Boolean
    Dim i%
    CheckData = True
    If m_sFlag = ID_ADDNEW Then
        With grdData
            For i = 1 To .Rows - 1
                If Trim(txtCode) = .TextMatrix(i, 1) Then
                    MsgBox LoadResString(114), vbInformation
                    txtCode.SetFocus
                    CheckData = False
                    Exit Function
                End If
            Next i
        End With
    End If
    
    If Len(txtName(0)) = 0 Then
        MsgBox "��ǰ���� �����ϴ�. ��ǰ���� �־� �ֽʽÿ�", vbInformation
        txtName(0).SetFocus
        CheckData = False
        Exit Function
    End If

End Function

Private Function SaveData() As Boolean
    Dim oArticle As PlusLib2.CArticle
    Dim NewArticle As PlusLib2.TArticle
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
        
    
    With NewArticle
        .sArticleID = IIf(Len(txtCode) > 0, Format(txtCode, "0000"), "")
        .sArticle = txtName(0)
        .sThreaID = txtName(1).Tag
        .sStuffWidthID = IIf(Len(txtName(2).Tag) > 0, txtName(2).Tag, "01")
        .DyeingID = cboDye.ListIndex
        .Weight = IIf(Len(txtName(3)) = 0, 0, txtName(3))
        
    End With
    
    Set oArticle = New PlusLib2.CArticle
    oArticle.Connection = adoCon
'    oArticle.UserName = g_sUserName
    
    Select Case m_sFlag
        Case ID_ADDNEW
            Set rs = oArticle.GetArticleByName(txtName(0))
            
            If Not rs.EOF Then
                MsgBox "������ �̸��� ǰ���� " & rs!ArticleID & " �� �ֽ��ϴ�"
                rs.Close
                Set rs = Nothing
                SaveData = True
            Else
                SaveData = oArticle.AddNewArticle(NewArticle)
            End If
            
        Case ID_UPDATE
            NewArticle.sArticleID = grdData.TextMatrix(grdData.Row, 1)
            SaveData = oArticle.UpdateArticle(NewArticle)
        Case ID_DELETE
            SaveData = oArticle.DeleteArticle(grdData.TextMatrix(grdData.Row, 1))
    End Select
    
    Set oArticle = Nothing
    Exit Function
ErrHandler:
    Set oArticle = Nothing

    Call ErrorBox(Err.Number, "Article.SaveData", Err.Description)
End Function

Private Sub cmdSelect_Click()
    If grdData.Rows = grdData.FixedRows Then Exit Sub

    Call SelectData
End Sub

Private Sub Form_Activate()
'    grdData.SetFocus
End Sub

Private Sub Form_Load()
    m_sFlag = ID_CANCEL
    
    Call SetOperate(Me)
    
    With cboDye
        .AddItem "Jigger"
        .AddItem "Rapid"
        .AddItem "CPB"
        
        .ListIndex = 0
    End With
    
    Call InitGrid
    
    txtCode.MaxLength = 5
    cmdAll.Picture = LoadResPicture("ALL", vbResIcon)
End Sub

Private Sub InitGrid()

    Call SetVSFlexGrid(grdData)
    With grdData
        .Redraw = False
        .Rows = 1
        .Cols = 9
        
        .TextArray(0) = ""
        .TextArray(1) = "�ڵ�":         .ColWidth(1) = 570:             .ColAlignment(1) = flexAlignCenterCenter
        .TextArray(2) = "ǰ��":         .ColWidth(2) = LIMIT_WIDTH:     .ColAlignment(2) = flexAlignLeftCenter
        .TextArray(3) = "��������":     .ColWidth(3) = 0
        .TextArray(4) = "��������ID":   .ColWidth(3) = 0
        .TextArray(5) = "������":       .ColWidth(4) = 0
        .TextArray(6) = "������ID":     .ColWidth(4) = 0
        .TextArray(7) = "������":       .ColWidth(5) = 0
        .TextArray(8) = "�߷�":         .ColWidth(6) = 0
        
        .Redraw = True
    End With
End Sub


Private Sub ChangeScroll()
    Dim lRows As Long
    
    lRows = GetVisibleVSGridRowCount(grdData)
    
    With grdData
        If lRows > LIMIT_ROW Then
            .ColWidth(2) = LIMIT_WIDTH - 240
        Else
            .ColWidth(2) = LIMIT_WIDTH
        End If
    End With

    If lRows = 0 Then
        Call ClearData
        cmdOperate(ID_UPDATE).Enabled = False
        cmdOperate(ID_DELETE).Enabled = False
    Else
        Call ShowData
        cmdOperate(ID_UPDATE).Enabled = True
        cmdOperate(ID_DELETE).Enabled = True
    End If
End Sub

Private Sub ClearData()
    Dim i%
    
    txtCode = ""
    
    For i = 0 To 3
        txtName(i) = ""
        txtName(i).Tag = ""
    Next i
    
End Sub

Private Sub ShowData()
    On Error Resume Next
    
    With grdData
        txtCode = .TextMatrix(.Row, 1)  '[�ڵ�]
        txtName(0) = .TextMatrix(.Row, 2) '[1] ǰ��
        txtName(1) = .TextMatrix(.Row, 3) '[2] ��������
        txtName(1).Tag = .TextMatrix(.Row, 4)   '���� ���� �ڵ�
        txtName(2) = .TextMatrix(.Row, 5)       ' ������
        txtName(2).Tag = .TextMatrix(.Row, 6)   ' ������ ���� �ڵ�
        cboDye.ListIndex = CLng(.TextMatrix(.Row, 7))       ' ������
        txtName(3) = .TextMatrix(.Row, 8)       ' �߷�
    End With
End Sub

Private Sub grdData_AfterSort(ByVal Col As Long, Order As Integer)
    Call ShowData
End Sub

Private Sub grdData_DblClick()
    With grdData
        If .MouseRow < .FixedRows Or .MouseRow >= .Rows Then Exit Sub
        
        Call SelectData
    End With
End Sub

Private Sub grdData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call SelectData
    End If
End Sub

Private Sub grdData_RowColChange()
    If m_bSkip Then Exit Sub
    Call ShowData
End Sub

Private Sub optSize_Click(Index As Integer)
    Dim lRows As Long
    
    lRows = GetVisibleVSGridRowCount(grdData)
    
    If optSize(0).Value Then '[0] ���
        With grdData
            .Width = 7625
            If lRows > LIMIT_ROW Then
                .ColWidth(2) = LIMIT_WIDTH + 1560
            Else
                .ColWidth(2) = LIMIT_WIDTH + 1800
            End If
            .ColWidth(3) = 2400
        End With
    Else '[1] ��
        With grdData
            .Width = 3420
            If lRows > LIMIT_ROW Then
                .ColWidth(2) = LIMIT_WIDTH - 240
            Else
                .ColWidth(2) = LIMIT_WIDTH
            End If
            .ColWidth(3) = 0
        End With
    End If
End Sub

Private Sub txtSearch_Change()
    Dim iLoop  As Integer
    Dim iCount As Integer
    Dim iNowRow As Integer

    On Error GoTo ErrHandler
    
    If Len(Trim(txtSearch)) > 0 Then
        With grdData
            .Redraw = False

            For iLoop = .FixedRows To .Rows - .FixedRows
                If InStr(UCase(.TextArray(iLoop * .Cols + 2)), UCase(txtSearch)) = 0 Then
                    .RowHidden(iLoop) = True
                    iCount = iCount + 1
                Else
                    .RowHidden(iLoop) = False
                    iNowRow = iLoop
                End If
            Next iLoop

            If iNowRow > .FixedRows Then
                .Row = iNowRow
                
                .Col = .FixedCols
                .ColSel = .Cols - 1
            End If

            .Redraw = True
            .TopRow = .Row
        End With
    Else
        Call cmdAll_Click
    End If

    If iCount > 0 Then
        cmdAll.Visible = True
    Else
        cmdAll.Visible = False
    End If
    
    Exit Sub

ErrHandler:
    Call ErrorBox(Err.Number, "Article.txtSearch_Change", Err.Description)

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    'Call MoveFocus(KeyCode)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        grdData.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        Call SetGrid(FL_BY_NAME, txtSearch)
    End If
End Sub




