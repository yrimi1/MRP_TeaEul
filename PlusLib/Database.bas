Attribute VB_Name = "Database"
Option Explicit

Public g_adoCon As ADODB.Connection

Public g_sUserName As String

'****************************************************************
'* Date: 2002-08-13
'*
'* Description:
'*  오류발생시 Error Log를 만든다 (SQL)
'****************************************************************
Public Sub ErrLogService(sErrLog() As String, nErrNO As Long, sErrMsg As String, Optional nErrIndex As Integer = 0)
    Dim adoCmd As ADODB.Command
    Dim i%, nErrID&

    On Error Resume Next

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "WizLog.dbo.xp_iErrLog"

        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 4, nErrID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, GetComputer())
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, g_sUserName)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nErrNO)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nErrIndex)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 500, sErrMsg)

        .Execute

        nErrID = .Parameters(0).Value

        .CommandText = "WizLog.dbo.xp_iErrLogSub"
        .Prepared = True

        For i = LBound(sErrLog) To UBound(sErrLog)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nErrID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, i)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1500, sErrLog(i))

            .Execute
        Next i
    End With
    Set adoCmd = Nothing
End Sub

'****************************************************************
'* Date: 2002-09-01
'*
'* Description:
'*  Logging
'****************************************************************
Public Sub LogService(sLog() As String)
    Dim adoCmd As ADODB.Command
    Dim i%, nLogID&

    On Error Resume Next

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "WizLog.dbo.xp_iLog"
        .Prepared = True
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, GetComputer())
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, g_sUserName)

        For i = LBound(sLog) To UBound(sLog)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1500, sLog(i))

            .Execute
            
            .Parameters.Delete (.Parameters.Count - 1)
        Next i
    End With
    Set adoCmd = Nothing
End Sub

Public Function GetMaxValue(sTable As String, sField As String, Optional sWhere As String = "") As Long
    Dim rs   As ADODB.Recordset
    Dim sSQL As String

    On Error GoTo ErrHandler

    Set rs = New ADODB.Recordset

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTable
    If Len(sWhere) > 0 Then
        If InStr(sWhere, "WHERE") Then
            sSQL = sSQL & " " & sWhere
        Else
            sSQL = sSQL & " WHERE " & sWhere
        End If
    End If

    rs.Open sSQL, g_adoCon, adOpenForwardOnly, adLockReadOnly

    If IsNull(rs(0)) Then
        GetMaxValue = 1
    Else
        GetMaxValue = CLng(rs(0)) + 1
    End If

    rs.Close
    Set rs = Nothing

    Exit Function

ErrHandler:
    Set rs = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'20091215 홍태영 추가 삼우에 코드번호 99번까지 다 차서 비어있는 값 임시로 씀.. StuffWidthID 3자리로 늘릴 필요 있음
Public Function GetMinValue(sTable As String, sField As String, Optional sWhere As String = "") As Long
    Dim rs   As ADODB.Recordset
    Dim sSQL As String

    On Error GoTo ErrHandler

    Set rs = New ADODB.Recordset
    
    sSQL = " SELECT MIN(idx) as minCnt FROM idx_Count" & vbCr
    sSQL = sSQL & " WHERE Idx not in (select StuffwidthID From mt_StuffWidth)"
    
    

    rs.Open sSQL, g_adoCon, adOpenForwardOnly, adLockReadOnly

    If IsNull(rs(0)) Then
        GetMinValue = 1
    Else
        GetMinValue = CLng(rs(0))
    End If

    rs.Close
    Set rs = Nothing

    Exit Function

ErrHandler:
    Set rs = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function IsExistRecord(sTable As String, Optional sWhere As String = "") As Boolean
    Dim rs   As ADODB.Recordset
    Dim sSQL As String

    On Error GoTo ErrHandler

    Set rs = New ADODB.Recordset

    sSQL = "SELECT * FROM " & sTable
    If Len(sWhere) > 0 Then
        If InStr(sWhere, "WHERE") Then
            sSQL = sSQL & " " & sWhere
        Else
            sSQL = sSQL & " WHERE " & sWhere
        End If
    End If

    rs.Open sSQL, g_adoCon, adOpenForwardOnly, adLockReadOnly

    IsExistRecord = IIf(rs.EOF, False, True)

    rs.Close
    Set rs = Nothing

    Exit Function

ErrHandler:
    Set rs = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'****************************************************************
'* Date: 2000-07-20 (FRI)
'*
'* Description:
'*  Command 객체의 Parameter 값을 초기화
'*
'****************************************************************
Public Sub ClearParameter(oCommand As ADODB.Command)
    Dim i%

    For i = oCommand.Parameters.Count - 1 To 0 Step -1
        oCommand.Parameters.Delete (i)
    Next i
End Sub

'****************************************************************
'*Author: Meridian I.S.
'*
'*Description:
'*  Parameter로 날아온 SQL문을 실행한다.
'*
'****************************************************************
Public Function HandleDB(SQL As String) As Boolean
    Dim sSQL(0) As String
    sSQL(0) = SQL

    On Error GoTo ErrHandle

    g_adoCon.Execute SQL, , adExecuteNoRecords

    Call LogService(sSQL)

    HandleDB = True

    Exit Function

ErrHandle:
    Call ErrLogService(sSQL, Err.Number, Err.Description, 0)
    HandleDB = False

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

'****************************************************************
'*Author: Meridian I.S.
'*
'*Description:
'*  Parameter로 날아온 다중 SQL문을 실행한다.
'*
'****************************************************************
Public Function HandleDBMulti(SQL() As String) As Boolean
    Dim iLoop As Integer

    On Error GoTo ErrHandle

    g_adoCon.BeginTrans
    For iLoop = LBound(SQL) To UBound(SQL)
        g_adoCon.Execute SQL(iLoop), , adExecuteNoRecords
    Next iLoop
    g_adoCon.CommitTrans

    Call LogService(SQL)

    HandleDBMulti = True

    Exit Function

ErrHandle:
    g_adoCon.RollbackTrans

    Call ErrLogService(SQL, Err.Number, Err.Description, iLoop)
    HandleDBMulti = False

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


