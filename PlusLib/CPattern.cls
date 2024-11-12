VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Type tPattern
    sPatternID      As String '[1] 공정 패턴 코드
    sPattern        As String '[2] 공정 패턴 명
    sWorkID         As String '[3] 가공 구분
End Type

Type TSubPattern
    sPatternID      As String    '[1] 공정 패턴 코드
    nPatternSeq     As Integer  '[2] 공정 패턴 순위
    sProcessID      As String    '[3] 공정 코드
End Type

Private Const CLASS_NAME As String = "CPattern"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetAllPattern() As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Pattern_sAllPattern"

        Set GetAllPattern = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

' 패턴  SELECT
Public Function GetPattern(Optional nChkWorkID As Integer = 0, Optional sWorkID As String) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Pattern_sPattern"
        '.CommandText = "sp_GetPattern"

        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkWorkID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sWorkID)
        Set GetPattern = .Execute
    End With
    
    Set adoCmd = Nothing
End Function
' 패턴 List
Public Function GetPatternList() As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PatternList"

        Set GetPatternList = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetPatternSub(NewKey As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Pattern_sPatternSub"
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, NewKey)
        Set GetPatternSub = .Execute
    End With
    
    Set adoCmd = Nothing
End Function


Public Function AddNewPattern(NewItem As tPattern, SubPattern() As TSubPattern, nSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim iLoop%, iCount%, nSql%
    Dim sPatternID As String
    Dim bError As Boolean
    Dim sLog() As String
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd     ' 새로운 공정 패턴 입력
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Pattern_iPattern"

        .Parameters.Append .CreateParameter(, adChar, adParamOutput, 2, NewItem.sPatternID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, NewItem.sPattern)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.sWorkID)

        .Execute
        
        sPatternID = .Parameters(0).Value
    End With
    
    nSql = 1
    
    If nSeq <> 0 Then
        With adoCmd     ' 공정패턴 내용 입력
            .CommandText = "xp_Pattern_iPatternSub"
            .Prepared = True
        
            For iLoop = 0 To nSeq
                Call ClearParameter(adoCmd)
        
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sPatternID)
                .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 4, SubPattern(iLoop).nPatternSeq)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, Format(SubPattern(iLoop).sProcessID, "0000"))
                
                .Execute
                
                nSql = nSql + 1
            Next iLoop
        End With
    End If

    g_adoCon.CommitTrans
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPattern.AddNewPattern"

LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(nSql - 1)
    With NewItem
        sLog(0) = "xp_Pattern_iPattern '" & sPatternID & "', " & .sPattern & ", '" & .sWorkID & "'"
    End With
    
    For iLoop = 1 To nSql - 1
        With SubPattern(iLoop - 1)
            sLog(iLoop) = "xp_Pattern_iPatternSub '" & sPatternID & "', " & vbTab & .nPatternSeq & vbTab & ", '" & .sProcessID & "'"
        End With
    Next iLoop

    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewPattern = False
''        Err.Raise Err.Number, "CPattern.AddNewPattern", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        AddNewPattern = True
    End If

End Function


Public Function UpdatePattern(NewItem As tPattern, SubPattern() As TSubPattern, nSeq As Integer) As Boolean

    Dim iLoop%, iCount%, nSql%
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd     ' 새로운 공정 패턴 입력
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Pattern_uPattern"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewItem.sPatternID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, NewItem.sPattern)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.sWorkID)

        .Execute
        
    End With
    
    nSql = 1
    
    If nSeq <> 0 Then
        With adoCmd     ' 공정패턴 내용 입력
            .CommandText = "xp_Pattern_iPatternSub"
            .Prepared = True
        
            For iLoop = 0 To nSeq
                Call ClearParameter(adoCmd)
        
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, NewItem.sPatternID)
                .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 4, SubPattern(iLoop).nPatternSeq)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, Format(SubPattern(iLoop).sProcessID, "0000"))
                
                .Execute
                
                nSql = nSql + 1
            Next iLoop
        End With
    End If

    g_adoCon.CommitTrans
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPattern.UpdatePattern"

LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(nSql - 1)
    With NewItem
        sLog(0) = "xp_Pattern_uPattern '" & .sPatternID & "', " & .sPattern & ", '" & .sWorkID & "'"
    End With
    
    For iLoop = 1 To nSql - 1
        With SubPattern(iLoop - 1)
            sLog(iLoop) = "xp_Pattern_iPatternSub '" & NewItem.sPatternID & "', " & vbTab & .nPatternSeq & vbTab & ", '" & .sProcessID & "'"
        End With
    Next iLoop

    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdatePattern = False
''        Err.Raise Err.Number, "Pattern.AddNewPattern", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        UpdatePattern = True
    End If


End Function


Public Function DeletePattern(NewKey As String) As Boolean
   Dim iLoop%, iCount%, nSql%
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd     ' 새로운 공정 패턴 입력
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Pattern_dPattern"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewKey)
    
        .Execute
        
    End With
    
    nSql = 1
    
    g_adoCon.CommitTrans
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPattern.DeletePattern"

LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)
    
    sLog(0) = "xp_Pattern_dPattern '" & NewKey & "'"
    
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeletePattern = False
''        Err.Raise Err.Number, "CPattern.DeletePattern", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        DeletePattern = True
    End If
    
End Function

