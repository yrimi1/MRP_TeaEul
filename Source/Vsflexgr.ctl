VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.UserControl WizFlexGroup 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   ForeColor       =   &H8000000E&
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   4965
   Begin VB.PictureBox picGroup 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   420
      Index           =   0
      Left            =   525
      ScaleHeight     =   420
      ScaleWidth      =   1260
      TabIndex        =   1
      Tag             =   "Hello"
      Top             =   360
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Align           =   2  '아래 맞춤
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   1065
      Width           =   4965
      _cx             =   8758
      _cy             =   6297
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorBkg    =   8421504
      BackColorAlternate=   -2147483643
      GridColor       =   14737632
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   12
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   6
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
      ExplorerBar     =   3
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
Attribute VB_Name = "WizFlexGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTONPUSH = &H10

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private Type POINTSGL
    x As Single
    y As Single
End Type

Private Type GROUPINFO
    ctl  As PictureBox
    Text As String
End Type

Private Const CLR_BTNFACE = &H8000000F
Private Const CLR_BTNSHADOW = &H80000010
Private Const CLR_BTNHILITE = &H80000014

Private Const HELPMSG = " 분류(Group)하고자 하는 컬럼(Column)의 헤더(Header)를 이곳에 끌어 놓으십시요. "
Private Const DRAG_TOLERANCE = 100 ' Twips

' Mouse control
Private m_bCapture  As Boolean  ' Mouse captured?
Private m_bDragging As Boolean  ' Dragging control?
Private m_ptDown    As POINTSGL ' Where was the click
Private m_ptControl As POINTSGL ' Original coordinates

Private m_iControl    As Integer
Private m_iGroups     As Integer    ' How many groups do we have
Private m_GroupInfo() As GROUPINFO  ' Group information vector

Private m_bLock()  As Boolean
Private m_bTotal() As Boolean
'이벤트 선언:
Event DblClick() 'MappingInfo=fg,fg,-1,DblClick

Public Property Get FlexGrid() As VSFlexGrid
    Set FlexGrid = fg
End Property

Public Property Let ColLock(Index As Integer, bFlag As Boolean)
    If fg.Cols - 1 = UBound(m_bLock) Then
        m_bLock(Index) = bFlag
    Else
        Dim i%
        ReDim Preserve m_bLock(fg.Cols - 1)

        For i = 0 To fg.Cols - 1
            m_bLock(i) = False
        Next i

        m_bLock(Index) = bFlag
    End If
End Property

Public Property Let ColTotal(Index As Integer, bFlag As Boolean)
    If fg.Cols - 1 = UBound(m_bTotal) Then
        m_bTotal(Index) = bFlag
    Else
        Dim i%
        ReDim Preserve m_bTotal(fg.Cols - 1)

        For i = 0 To fg.Cols - 1
            m_bTotal(i) = False
        Next i

        m_bTotal(Index) = bFlag
    End If
End Property

Public Sub Update()
    UpdateLayout True
End Sub

Private Sub UserControl_Resize()
    UpdateLayout False
End Sub

Private Sub picGroup_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Left button starts dragging
    If Button = 1 Then
        ' Save dragging information
        m_bCapture = True
        m_bDragging = False
        m_ptDown.x = x
        m_ptDown.y = y

        ' Bring control to top, save its original position
        With picGroup(Index)
            .ZOrder
            m_ptControl.x = .Left
            m_ptControl.y = .Top
        End With
    End If
End Sub

Private Sub picGroup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Drag control around
    If m_bCapture Then
        With picGroup(Index)
            ' If we are not dragging yet, maybe it's time to start
            If Not m_bDragging Then
                If Abs(x - m_ptDown.x) > DRAG_TOLERANCE Then m_bDragging = True
                If Abs(y - m_ptDown.y) > DRAG_TOLERANCE Then m_bDragging = True
            End If

            ' If we're dragging, then do it
            If m_bDragging Then
                ' Get new coordinates
                x = .Left + (x - m_ptDown.x)
                y = .Top + (y - m_ptDown.y)

                ' Restrict boundaries
                If x < 0 Then x = 0
                If y < 0 Then y = 0
                If x > UserControl.ScaleWidth - .Width Then x = UserControl.ScaleWidth - .Width
                If y > UserControl.ScaleHeight - .Height Then y = UserControl.ScaleHeight - .Height
                If y > fg.Top Then y = fg.Top

                ' Move the control
                .Move x, y

                ' Show where we'd go if we dropped now
                ' UNDONE
            End If
        End With
    End If
End Sub

Private Sub picGroup_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If we were dragging, we may have just moved the group to a new position, or we may have dropped it back into the grid
    If m_bDragging Then
        With fg
            .Redraw = False

            ' Back into grid, different position
            y = picGroup(Index).Top + y
            If y > .Top Then
                ' See which column it was and where the mouse is
                Dim iCol1%, iCol2%

                iCol1 = FindColumn(m_GroupInfo(picGroup(Index).Tag).Text)
                iCol2 = IIf(.MouseCol < 0, 0, .MouseCol)
    
                ' Different? move column
                If iCol2 <> iCol1 Then
                    If Not m_bLock(iCol2) Then .ColPosition(iCol1) = iCol2
                ' Same? switch sort order
                Else
                    If .ColSort(iCol2) = flexSortGenericAscending Then
                        .ColSort(iCol2) = flexSortGenericDescending
                    Else
                        .ColSort(iCol2) = flexSortGenericAscending
                    End If
                End If

                ' Remove our brand-new group
                DeleteGroup Index
            End If

            ' Either way, show changes
            UpdateLayout True

            .Redraw = True
        End With
    End If

    ' Cancel capture no matter what
    m_bCapture = False

    Call CalcTotalAmnt
End Sub

Private Sub picGroup_Click(Index As Integer)
    ' Unless we were dragging, Revert sort direction
    If (Not m_bDragging) And (m_ptControl.x > -1) Then
        ' Revert sort direction
        Dim iIdx%

        iIdx = CInt(picGroup(Index).Tag)
        If fg.ColSort(iIdx) = flexSortGenericDescending Then
            fg.ColSort(iIdx) = flexSortGenericAscending
        Else
            fg.ColSort(iIdx) = flexSortGenericDescending
        End If

        ' Show the change
        UpdateLayout True
    End If
End Sub

Private Sub picGroup_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Escape cancels dragging/clicking
    If (KeyAscii = 27) And (m_bCapture = True) Then
        ' Move control back to its original position
        If m_bDragging Then
            ' If the group was still being created (not just dragged), delete it
            If m_ptControl.x < 0 And m_ptControl.y < 0 Then
                DeleteGroup Index
            ' Otherwise, move it back to where it was
            Else
                picGroup(Index).Move m_ptControl.x, m_ptControl.y
            End If
        End If

        ' Reset state variables
        m_bCapture = False
        m_bDragging = True
    End If
End Sub

Private Sub picGroup_Paint(Index As Integer)
    Dim rc As RECT

    With picGroup(Index)
        ' Draw frame
        rc.Top = 0
        rc.Left = 0
        rc.Right = .Width / Screen.TwipsPerPixelX
        rc.Bottom = .Height / Screen.TwipsPerPixelY
        DrawFrameControl .hDC, rc, DFC_BUTTON, DFCS_BUTTONPUSH

        ' Draw text
        .CurrentX = .TextWidth(" ")
        .CurrentY = (.Height - .TextHeight(" ")) / 2.5
        picGroup(Index).Print m_GroupInfo(.Tag).Text

        ' Draw sort arrow if this is a group already
        If fg.ColWidth(.Tag) = 0 Then
            Dim x As Single, y As Single, sz As Single
            sz = .Height * (1 / 3)
            x = .Width - sz

            ' Pointing up
            If fg.ColSort(.Tag) = flexSortGenericDescending Then
                y = (.Height - sz) / 2 + sz
                picGroup(Index).Line (x, y)-(x - sz, y), CLR_BTNHILITE
                picGroup(Index).Line -(x - sz / 2, y - sz), CLR_BTNSHADOW
                picGroup(Index).Line -(x, y), CLR_BTNHILITE
            ' Pointing down
            Else
                y = (.Height - sz) / 2
                picGroup(Index).Line (x, y)-(x - sz, y), CLR_BTNSHADOW
                picGroup(Index).Line -(x - sz / 2, y + sz), CLR_BTNSHADOW
                picGroup(Index).Line -(x, y), CLR_BTNHILITE
            End If
        End If
    End With
End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    ' If we clicked on a column, start dragging it
    If Button = 1 And Shift = 0 And fg.MouseRow = 0 Then
        ' Make sure we don't group on everything
        If m_iGroups >= fg.Cols - 1 Then Exit Sub

        ' Which column are we grouping on?
        Dim Col%
        Col = IIf(fg.MouseCol < 0, 0, fg.MouseCol)

        If m_bLock(Col) Then Exit Sub

        Dim i%

        ' Confirm that this is a groupable column
        For i = 0 To m_iGroups - 1
            If m_GroupInfo(i).Text = fg.Cell(flexcpTextDisplay, 0, Col) Then
                Cancel = True
                Beep
                Exit Sub
            End If
        Next
        ' UNDONE

        ' Create entry in global array
        i = m_iGroups
        m_iGroups = m_iGroups + 1
        ReDim Preserve m_GroupInfo(i)

        ' Create new group control
        m_iControl = m_iControl + 1
        Load picGroup(m_iControl)
        Set m_GroupInfo(i).ctl = picGroup(m_iControl)
        m_GroupInfo(i).Text = fg.Cell(flexcpTextDisplay, 0, Col)

        ' Init group control
        With picGroup(m_iControl)
            .Tag = i
            .Width = .TextWidth(m_GroupInfo(i).Text) + 2 * fg.RowHeight(0)
            .Height = fg.RowHeight(0) * 1.1
            .Move fg.ColPos(Col), fg.Top
            .Font = fg.Font
            .ZOrder
        End With

        ' Save original position (none in this case)
        m_ptControl.x = -1
        m_ptControl.y = -1

        ' Start dragging
        m_bCapture = True
        m_bDragging = True
        m_ptDown.x = x - picGroup(m_iControl).Left
        m_ptDown.y = fg.Top + y - picGroup(m_iControl).Top
        picGroup_Paint m_iControl

        ' This is really cool:
        '   Flex got the mouse down, but we want the group control to handle it
        '   so we set Cancel to true and transfer the mouse to the group control
        '   using the SetCapture API.
        Cancel = True

        With picGroup(m_iControl)
            .Visible = True
            .SetFocus
            SetCapture .hwnd
        End With
    End If
End Sub

Private Sub fg_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If m_bLock(Col) Then Position = Col
End Sub

Private Sub UserControl_Initialize()
    ' Initialize embedded FlexGrid
    With fg
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .OutlineBar = flexOutlineBarComplete
        .ExplorerBar = flexExSortAndMove
    End With

    ' Initialize group control based on grid data
    With picGroup(0)
        .Font = fg.Font
        .Height = fg.RowHeight(0)
        .Tag = 0
    End With

    ReDim m_bLock(0)
    ReDim m_bTotal(0)
End Sub

Private Sub UpdateLayout(dogrid As Boolean)
    Dim swap As GROUPINFO
    Dim i%, cnt%, Done%
    Dim x As Single, y As Single, rh As Single
    Dim offsety As Single

    ' See how many groups are visible
    cnt = m_iGroups

    ' Dimension and clear grouping area
    rh = fg.RowHeight(0)
    offsety = rh / 2
    y = 2 * fg.RowHeight(0)
    If cnt > 1 Then y = y + (cnt - 1) * offsety
    y = UserControl.ScaleHeight - y
    If y < 0 Then y = 0
    fg.Height = y
    UserControl.Cls

    ' If no groups, show helpful message
    If cnt = 0 Then
        UserControl.CurrentX = rh / 2
        UserControl.CurrentY = rh / 2
        UserControl.Print HELPMSG
    End If

    ' Sort group vector by position (left-to-right)
    While Not Done
        Done = True
        For i = 0 To cnt - 2
            If m_GroupInfo(i).ctl.Left > m_GroupInfo(i + 1).ctl.Left Then
                Done = False
                swap = m_GroupInfo(i)
                m_GroupInfo(i) = m_GroupInfo(i + 1)
                m_GroupInfo(i + 1) = swap
            End If
        Next
    Wend

    ' Each control gets and index into the vector
    For i = 0 To cnt - 1
        m_GroupInfo(i).ctl.Tag = i
    Next

    ' Position group controls
    y = rh / 2
    x = y
    For i = 0 To cnt - 1
        With m_GroupInfo(i).ctl
            ' Move the control
            .Move x, y
            y = y + offsety
            x = x + .Width + rh / 3

            ' Draw connector
            If i < cnt - 1 Then
                UserControl.Line (x, y + 2 / 3 * rh)-(x - rh * 2 / 3, y + 2 / 3 * rh), 0
                UserControl.Line -(x - rh * 2 / 3, y + rh / 2 - Screen.TwipsPerPixelY), 0
            End If

            ' Draw placeholder
            UserControl.Line (.Left, .Top)-(.Left + .Width - Screen.TwipsPerPixelX, .Top + .Height - Screen.TwipsPerPixelY), 0, B
        End With
    Next

    ' Redraw all controls at their new positions
    For i = 0 To cnt - 1
        picGroup_Paint m_GroupInfo(i).ctl.Index
    Next
    UserControl.Refresh

    ' Update the grid
    If dogrid Then UpdateGrid

    ' Redraw all controls at their new positions (to show sort direction)
    For i = 0 To cnt - 1
        picGroup_Paint m_GroupInfo(i).ctl.Index
    Next
End Sub

Private Sub UpdateGrid()
    ' Redraw is off to speed things up
    fg.Redraw = False

    ' Move groups to left
    Dim i%, Col%

    For i = 0 To m_iGroups - 1
        Col = FindColumn(m_GroupInfo(i).Text)
        fg.ColPosition(Col) = i
    Next

    ' Hide groups, Make sure they're all sortable
    For i = 0 To m_iGroups - 1
        fg.ColHidden(i) = True
        If fg.ColSort(i) = 0 Then fg.ColSort(i) = flexSortGenericAscending
    Next

    ' Show non-groups
    For i = m_iGroups To fg.Cols - 1
        fg.ColHidden(i) = False
    Next

    ' Sort
    If fg.Rows > 1 Then
        fg.Select fg.Row, 0, fg.Row, fg.Cols - 1
        fg.Sort = flexSortUseColSort
    End If

    ' Create groups
    fg.Subtotal flexSTClear
    
    If m_iGroups > 0 Then
        With fg
            For i = 0 To m_iGroups - 1
                .Subtotal flexSTNone, i, , , CLR_BTNFACE, , , , , True
            Next

            ' Group them
            .Outline m_iGroups - 1
            .OutlineCol = m_iGroups
            .AutoSize m_iGroups
        End With
    End If

    ' Move text to visible rows
    Dim s$
    If m_iGroups > 0 Then
        With fg
            For i = 1 To .Rows - 1
                If .IsSubtotal(i) Then
                    s = .Cell(flexcpTextDisplay, i, 0)
                    .Cell(flexcpText, i, 0) = ""
                    .Cell(flexcpText, i, m_iGroups) = s
                    .Cell(flexcpAlignment, i, m_iGroups) = flexAlignLeftCenter
                End If
            Next i
        End With
    End If

    fg.MergeCells = flexMergeSpill

    ' Redraw is back on
    fg.Redraw = True
End Sub

' Remove control from the list
Private Sub DeleteGroup(Index As Integer)
    Dim i%

    For i = CLng(picGroup(Index).Tag) To m_iGroups - 2
        m_GroupInfo(i) = m_GroupInfo(i + 1)
    Next
    m_iGroups = m_iGroups - 1

    If m_iGroups = 0 Then fg.Outline 1

    ' Hide/unload the control
    picGroup(Index).Visible = False
    If Index > 0 Then Unload picGroup(Index)
End Sub

Private Function FindColumn(s$) As Integer
    ' Locate column based on header text
    Dim i%

    For i = 0 To fg.Cols - 1
        If fg.Cell(flexcpTextDisplay, 0, i) = s Then
            FindColumn = i
            Exit Function
        End If
    Next

    ' This should never happen
    FindColumn = -1
End Function

Private Sub CalcTotalAmnt()
    Dim i%, j%, k%, iRow%()
    Dim iTemp As Currency
    Dim iCol%, nCheck%
    On Error Resume Next

    With fg
        ReDim iRow(m_iGroups - 1)
        
        ' iRow() Initialize
        For i = 0 To m_iGroups
            iRow(i) = -1
        Next i

        ' Search Count Column
        For i = 0 To .Cols - 1
            If m_bTotal(i) Then
                iCol = i - 1
                Exit For
            End If
        Next i


        For i = 1 To .Rows - 1 ' i : Row
            If .IsSubtotal(i) Then ' Special Row (Group Row)
                .TextMatrix(i, iCol) = "0"
                If iRow(.RowOutlineLevel(i)) <> i Then
                    .TextMatrix(iRow(.RowOutlineLevel(i)), iCol) = .TextMatrix(iRow(.RowOutlineLevel(i)), iCol) & " 건"
                    iRow(.RowOutlineLevel(i)) = i
                End If
            Else        ' General Row
                For j = 0 To .Cols - 1 ' j : Column
                    If m_bTotal(j) And IsNumeric(.TextMatrix(i, j)) Then ' Sum Column
                        nCheck = nCheck + 1 ' Sum Column 여러개일 경우 중복 제거
                        For k = 0 To m_iGroups ' k : Level Loop
                            If Not IsNumeric(.TextMatrix(iRow(k), j)) Then
                                iTemp = 0
                            Else
                                iTemp = CCur(.TextMatrix(iRow(k), j))
                            End If
                            .TextMatrix(iRow(k), j) = SetCurrency(iTemp + CCur(.TextMatrix(i, j)))
                            If nCheck = 1 Then
                                .TextMatrix(iRow(k), iCol) = CCur(.TextMatrix(iRow(k), iCol)) + 1
                            End If
                        Next k
                    End If
                Next j
                nCheck = 0
            End If
        Next i
        
        For i = 0 To m_iGroups
            .TextMatrix(iRow(i), iCol) = .TextMatrix(iRow(i), iCol) & " 건"
        Next i
    End With
End Sub

Public Sub InitGroup()
    Dim i%

    For i = 0 To m_iGroups - 1
        m_GroupInfo(i).ctl.Visible = False
        Unload m_GroupInfo(i).ctl
    Next i

    ReDim m_GroupInfo(0)
    m_iControl = 0
    m_iGroups = 0

    fg.Outline 1
End Sub
Private Sub fg_DblClick()
    RaiseEvent DblClick
End Sub

