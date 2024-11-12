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
    sRoundClss       As String   '소숫점관리
    sGradeClss       As String   '등급결정방법
    sDemeritClss     As String   '벌점적용방법
    sLossClss        As String   '보상적용방법
    sDefectClss  As String   '대표불량등급
    sCutDefect   As String   '난단대표불량사용여부
    nButtonX     As Integer   '불량버튼 개수 X
    nButtonY     As Integer   '불량버튼 개수 Y
    nColorCnt    As Integer   '사용색상수
    nRepeatCnt   As Integer   '반복색상수
    nFontSize    As Integer   '불량명 폰트 크기
    nBasePosition As Single  '시작위치조정
    nRollClss       As Integer '절번호 관리 0 : Order 별 1 : Color 별 2 : Order Color Lot별 3 : Order Lot별 4 : Order, Color 호기별
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
