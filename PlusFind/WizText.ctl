VERSION 5.00
Begin VB.UserControl WizText 
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   ScaleHeight     =   1635
   ScaleWidth      =   4095
   Begin VB.TextBox TextBox 
      BackColor       =   &H8000000E&
      Height          =   780
      Left            =   405
      TabIndex        =   0
      Top             =   330
      Width           =   3315
   End
End
Attribute VB_Name = "WizText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'이벤트 선언:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TextBox,TextBox,-1,KeyDown
Attribute KeyDown.VB_Description = "개체에 포커스가 있을 때 키를 누르면 발생합니다."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=TextBox,TextBox,-1,KeyPress
Attribute KeyPress.VB_Description = "ANSI키를 누르고 놓았을 경우 발생합니다."
Event Change() 'MappingInfo=TextBox,TextBox,-1,Change
Attribute Change.VB_Description = "컨트롤의 내용이 변경될 때 발생합니다."


Private Sub TextBox_GotFocus()
    With TextBox
        .BackColor = vbCyan ' vbInactiveTitleBar 'vbHighlight
        .ForeColor = vbWindowText
        
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TextBox_LostFocus()
    With TextBox
        .BackColor = vbWindowBackground
        .ForeColor = vbWindowText
    End With
End Sub

Private Sub TextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyDown Then
        SendKeys "{TAB}"
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TextBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
    RaiseEvent KeyPress(KeyAscii)
End Sub
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "사용자가 만든 이벤트에 대해 개체가 응답할 수 있는지의 여부를 결정하는 값을 반환하거나 설정합니다."
    Enabled = TextBox.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    TextBox.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font 개체를 반환합니다."
Attribute Font.VB_UserMemId = -512
    Set Font = TextBox.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set TextBox.Font = New_Font
    PropertyChanged "Font"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,BorderStyle
Public Property Get BorderStyle() As MSComctlLib.BorderStyleConstants
Attribute BorderStyle.VB_Description = "개체 테두리 유형을 반환하거나 설정합니다."
    BorderStyle = TextBox.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MSComctlLib.BorderStyleConstants)
    TextBox.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "개체를 완전히 다시 그리게 합니다."
    UserControl.Refresh
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "CheckBox나 OptionButton 또는  컨트롤 텍스트의 맞춤을 반환하거나 설정합니다."
    Alignment = TextBox.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    TextBox.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "컨트롤의 편집 가능 여부를 결정합니다."
    Locked = TextBox.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    TextBox.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "컨트롤에 들어갈 수 있는 문자의 최대수를 반환하거나 설정합니다."
    MaxLength = TextBox.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    TextBox.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "컨트롤에 포함된 텍스트를 반환하거나 설정합니다."
Attribute Text.VB_UserMemId = 0
    Text = TextBox.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    TextBox.Text() = New_Text
    PropertyChanged "Text"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "컨트롤이 여러 줄의 텍스트를 받아들일 수 있는지 여부를 결정하는 값을 반환하거나 설정합니다."
    MultiLine = TextBox.MultiLine
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,ScrollBars
Public Property Get ScrollBars() As ScrollBarConstants
Attribute ScrollBars.VB_Description = "개체가 수직/수평 스크롤 막대를 가지는지의 여부를 나타내는 값을 반환하거나 설정합니다."
    ScrollBars = TextBox.ScrollBars
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "선택된 텍스트의 시작점을 반환하거나 설정합니다."
    SelStart = TextBox.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    TextBox.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "현재 선택된 텍스트를 포함하는 문자열을 반환하거나 설정합니다."
    SelText = TextBox.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    TextBox.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

Private Sub TextBox_Change()
    RaiseEvent Change
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "마우스가 컨트롤에 일시 정지되어 있을 때 표시되는 텍스트를 반환하거나 설정합니다."
    ToolTipText = TextBox.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    TextBox.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    TextBox.Enabled = PropBag.ReadProperty("Enabled", True)
    Set TextBox.Font = PropBag.ReadProperty("Font", Ambient.Font)
    TextBox.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    TextBox.Alignment = PropBag.ReadProperty("Alignment", 0)
    TextBox.Locked = PropBag.ReadProperty("Locked", False)
    TextBox.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    TextBox.Text = PropBag.ReadProperty("Text", "")
    TextBox.SelStart = PropBag.ReadProperty("SelStart", 0)
    TextBox.SelText = PropBag.ReadProperty("SelText", "")
    TextBox.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    TextBox.BackColor = PropBag.ReadProperty("BackColor", &H8000000E)
    TextBox.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
End Sub

Private Sub UserControl_Resize()
    TextBox.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", TextBox.Enabled, True)
    Call PropBag.WriteProperty("Font", TextBox.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", TextBox.BorderStyle, 1)
    Call PropBag.WriteProperty("Alignment", TextBox.Alignment, 0)
    Call PropBag.WriteProperty("Locked", TextBox.Locked, False)
    Call PropBag.WriteProperty("MaxLength", TextBox.MaxLength, 0)
    Call PropBag.WriteProperty("Text", TextBox.Text, "")
    Call PropBag.WriteProperty("SelStart", TextBox.SelStart, 0)
    Call PropBag.WriteProperty("SelText", TextBox.SelText, "")
    Call PropBag.WriteProperty("ToolTipText", TextBox.ToolTipText, "")
    Call PropBag.WriteProperty("BackColor", TextBox.BackColor, &H8000000E)
    Call PropBag.WriteProperty("ForeColor", TextBox.ForeColor, &H80000008)
End Sub


'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "개체의 텍스트나 그래픽을 표시하기 위해 사용되는 배경색을 반환하거나 설정합니다."
    BackColor = TextBox.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    TextBox.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=TextBox,TextBox,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "개체에서 텍스트나 그래픽을 표시하는 전경색을 반환하거나 설정합니다."
    ForeColor = TextBox.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    TextBox.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

