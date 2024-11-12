VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Const CLASS_NAME As String = "CResult"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub



Public Function UpdateProecssResult(TResult As TProcessResult) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog(0) As String, bError As Boolean
    Dim i%, iSql%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Insert Inspect
        .CommandText = "xp_Work_uResultByProcess"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 9, TResult.sOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, TResult.sCardID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TResult.sColorID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, , TResult.nWorkSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TResult.sProcessID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, TResult.sMachineID)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TResult.nReduceRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 2, TResult.nNormalDensity)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TResult.nAPRTemper)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TResult.nSTMTemper)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TResult.nVelocity)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TResult.nTemper)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TResult.nDensity)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TResult.nWidth)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TResult.nFlexRate1)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, TResult.nFlexRate2)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TResult.nMedicine1)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TResult.nMedicine2)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TResult.nMedicine3)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TResult.nMedicine4)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TResult.nMedicine5)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TResult.sWorkDate)
                        
        .Execute
        iSql = 1
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True

LogMessage:
    Set adoCmd = Nothing

    With TResult
        sLog(0) = "xp_Work_uResult '" & .sOrderID & "', '" & .sCardID & "', " & .sColorID & ", '" & .nWorkSeq & "', '" & .sProcessID & "', '" & _
            .sMachineID & "', '" & .nReduceRate & ", " & .nNormalDensity & ", " & .nAPRTemper & ", " & .nSTMTemper & ", " & _
            .nVelocity & ", " & .nTemper & ", " & .nDensity & ", " & .nWidth & ", " & .nFlexRate1 & ", " & .nFlexRate2 & ", " & _
            .nMedicine1 & ", " & .nMedicine2 & ", " & .nMedicine3 & ", " & .nMedicine4 & ", " & .nMedicine5 & ", '" & .sWorkDate & "'"
    End With
    

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSql)

        UpdateProecssResult = False
        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateProecssResult"), Err.Description
    Else
        Call LogService(sLog)

        UpdateProecssResult = True
    End If


End Function



Public Function GetResultByProcess(sSDate As String, sEDate As String, sProcessID As String, nChkMachineID As Integer, _
    sMachineID As String, nChkTeamID As Integer, sTeamID As String, nChkOrder As Integer, sOrder As String, _
    nChkCustomID As Integer, sCustomID As String, nChkArticleID As Integer, sArticleID As String) As Recordset
    
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    ' ÅÙÅÍ, °ÇÁ¶, ¼ö¼¼, ¸ð¼Ò, m/c, cpbÀüÃ³¸®,  peach, ¼¤Çª
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        Select Case CInt(sProcessID)
            Case 4      ' ÅÙÅÍ
                .CommandText = "xp_Work_sResultWithTenter"
            Case 5      ' °ÇÁ¶
                .CommandText = "xp_Work_sResultWithDry"
            Case 6   ' ¼ö¼¼
                .CommandText = "xp_Work_sResultWithRefine"
            Case 9   ' ¸ð¼Ò
                .CommandText = "xp_Work_sResultWithMoso"
            Case 7      ' m/c
                .CommandText = "xp_Work_sResultWithMercerizing"
            Case 8      ' cpb ÀüÃ³¸®
                .CommandText = "xp_Work_sResultWithCPBPre"
            Case 10     ' peach
                .CommandText = "xp_Work_sResultWithPeach"
            Case Else   ' ¼¤Çª
                .CommandText = "xp_Work_sResultWithShampu"
        End Select

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sProcessID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkMachineID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sMachineID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkTeamID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sTeamID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 15, sOrder)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        
        Set GetResultByProcess = .Execute
    End With

    Set adoCmd = Nothing
End Function


Public Function DeleteResult(sProcessID As String, sOrderID As String, sCardID As String, nWorkSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    bError = False
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        .CommandText = "xp_Card_dWorkResult"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sProcessID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 9, sOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sCardID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nWorkSeq)
        
        .Execute
    End With
    
    Set adoCmd = Nothing

    GoTo LogMessage

ErrHandler:
    bError = True
    
LogMessage:
    ReDim sLog(0)
    
    sLog(0) = "xp_Card_dWorkResult '" & sProcessID & "', '" & sOrderID & "', '" & sCardID & "', " & nWorkSeq & " "
        
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, 0)
        
        DeleteResult = False
        Err.Raise Err.Number, CLASS_NAME & ".DeleteResult", Err.Description
    Else
        Call LogService(sLog)
        
        DeleteResult = True
    End If
   
End Function



