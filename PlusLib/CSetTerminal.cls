VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSetTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type TSetTerminal
    sRoundClss       As String   '�Ҽ�������
    sGradeClss       As String   '��ް������
    sDemeritClss     As String   '����������
    sLossClss        As String   '����������
    sDefectClss  As String   '��ǥ�ҷ����
    sCutDefect   As String   '���ܴ�ǥ�ҷ���뿩��
    nButtonX     As Integer   '�ҷ���ư ���� X
    nButtonY     As Integer   '�ҷ���ư ���� Y
    nColorCnt    As Integer   '�������
    nRepeatCnt   As Integer   '�ݺ������
    nFontSize    As Integer   '�ҷ��� ��Ʈ ũ��
    nBasePosition As Single  '������ġ����
    nRollClss       As Integer '����ȣ ���� 0 : Order �� 1 : Color �� 2 : Order Color Lot�� 3 : Order Lot�� 4 : Order, Color ȣ�⺰
End Type

Private Const CLASS_NAME As String = "CSetTerminal"

Public Property Let Connection(g_adoConnect As Object)
    Set g_adoCon = g_adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Public Function GetSetTerminal() As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WizTerm_sTerminal"

        Set GetSetTerminal = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetTerminalDefect() As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WizIns_sDefectsub"

        Set GetTerminalDefect = .Execute
    End With

    Set adoCmd = Nothing
End Function


Public Function AddNewSetTerminal(TTerminal As TSetTerminal) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    
    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    
    bError = False
    
    g_adoCon.BeginTrans
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WizTerm_iSetTerminal"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, TTerminal.sRoundClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 9, TTerminal.sGradeClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 5, TTerminal.sDemeritClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 5, TTerminal.sLossClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 5, TTerminal.sDefectClss)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 5, TTerminal.sCutDefect)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, TTerminal.nButtonX)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, TTerminal.nButtonY)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, TTerminal.nColorCnt)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, TTerminal.nRepeatCnt)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, TTerminal.nFontSize)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, TTerminal.nRollClss)
    
        .Execute
    End With
    g_adoCon.CommitTrans
    
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CSetTerminal.AddNewSetTerminal"
LogMessage:
    Set adoCmd = Nothing
    ReDim sLog(0)
    
    With TTerminal
        sLog(0) = "xp_Term_iTerminal '" & .sRoundClss & "', '" & .sGradeClss & "', '" & .sDemeritClss & "', '" & .sLossClss & "', '" & .sDefectClss & "', '" & _
                .sCutDefect & "', " & .nButtonX & ", " & .nButtonY & ", " & .nColorCnt & ", " & .nRepeatCnt & ", " & _
                 .nFontSize & ", " & .nRollClss & " "
    End With

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, 0)

        AddNewSetTerminal = False
''        Err.Raise Err.Number, CLASS_NAME & ".AddNewSetTerminal", Err.Description
    Else
        Call LogService(sLog)

        AddNewSetTerminal = True
    End If
End Function