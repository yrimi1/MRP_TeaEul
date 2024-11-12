VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ECODE
    CD_DEPART   ' 부서관리
    CD_DUTY     ' 직책 관리
    CD_TRADE    ' 거래기준 관리
    CD_BASIS    ' 검사기준 관리
    CD_TEAM     ' 작업조 관리
    CD_WIDTH    ' 원단폭 관리
    CD_WORK     ' 가공구분 관리
    CD_LABEL    ' 레벨구분 관리
    CD_BAND     ' 밴드구분 관리
    CD_TAG      ' 태그구분 관리
    CD_FORM     ' 주문형태 관리
    CD_CLASS    ' 주문구분 관리
    CD_KIND     ' 불량구분 관리
    CD_OUTCLSS  ' 출고구분 관리
    CD_BACKCLSS ' 반품구분 관리
    CD_GRADE    ' 등급 관리
    CD_PROCESS  ' 공정코드
    CD_MACHINE  ' 기계코드
    CD_DEFECT   ' 불량코드
    CD_LENGTH   ' 필장
    CD_EXWORK   ' 외주가공
    CD_STANDARD ' 규격
    CD_THREAD   ' 사종구분(실의종류)
    CD_BREAK    ' 정대구분
    CD_COLOR    ' 색상구분
    CD_DRY      ' 건조정도
    CD_WaterRate    ' 액비
    CD_BatJa    ' 실물 밧자번호
End Enum

Private Const CLASS_NAME As String = "CCode"

Public Type TCode
    sCodeID As String   ' 코드
    sCode   As String   ' 명칭
End Type

Private m_sName As String

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Property Let CodeType(NewValue As ECODE)
    Select Case NewValue
        Case CD_DEPART: m_sName = "Depart"
        Case CD_DUTY: m_sName = "Duty"
        Case CD_TRADE: m_sName = "Trade"
        Case CD_BASIS: m_sName = "Basis"
        Case CD_TEAM: m_sName = "Team"
        Case CD_WIDTH: m_sName = "StuffWidth"
        Case CD_WORK: m_sName = "Work"
        Case CD_LABEL: m_sName = "Label"
        Case CD_BAND: m_sName = "Band"
        Case CD_TAG: m_sName = "Tag"
        Case CD_FORM: m_sName = "Form"
        Case CD_CLASS: m_sName = "Class"
        Case CD_KIND: m_sName = "Kind"
        Case CD_OUTCLSS: m_sName = "Outware"  '''기존 outclss 에서 outware로 변경.
        Case CD_BACKCLSS: m_sName = "Back"   ''기존 backclss 에서 back으로 변경...
        Case CD_GRADE: m_sName = "Grade"
        Case CD_PROCESS: m_sName = "Process"
        Case CD_MACHINE: m_sName = "Machine"
        Case CD_DEFECT: m_sName = "Defect"
        Case CD_LENGTH: m_sName = "Length"
        Case CD_EXWORK: m_sName = "ExWork"
        Case CD_STANDARD: m_sName = "Standard"
        Case CD_THREAD: m_sName = "Thread"
        Case CD_BREAK: m_sName = "Break"
        Case CD_COLOR: m_sName = "Color"
        Case CD_DRY: m_sName = "Dry"
        Case CD_WaterRate:  m_sName = "WaterRate"
        Case CD_BatJa:  m_sName = "BatJa"
    End Select
End Property

Public Function GetStuffWidth() As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sStuffWidthByWidth"
        
        Set GetStuffWidth = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetStuffWidth"), Err.Description
End Function

Public Function GetCode(Optional sBasisID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_s" & m_sName
        
        If m_sName = "Defect" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sBasisID)
        End If

        Set GetCode = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetCode"), Err.Description
End Function

Public Function GetCodeOne(Optional sCodeName) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    If IsMissing(sCodeName) Then
        sCodeName = ""
    Else
        sCodeName = Trim(sCodeName)
    End If

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_s" & m_sName & "One"

        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 35, sCodeName)

        Set GetCodeOne = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetCodeOne"), Err.Description
End Function

Public Function GetCodeID(sCodeID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_s" & m_sName & "ID"

        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 2, sCodeID)

        Set GetCodeID = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetCodeID"), Err.Description
End Function

Public Function GetMachineCode(sProcessID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sMachineID"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sProcessID)

        Set GetMachineCode = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetMachineCode"), Err.Description
End Function

Public Function GetMachine(Optional sProcessID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sMachine"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sProcessID)
          
        Set GetMachine = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetMachine"), Err.Description
End Function

Public Function GetDefectByLang(nChkLang As Integer, sBasisID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sDefectByLang"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkLang)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sBasisID)

        Set GetDefectByLang = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDefectByLang"), Err.Description
End Function

Public Function AddNewCode(tData As TCode) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%

    On Error GoTo ErrHandler
    
    With tData
        If Len(RTrim(.sCodeID)) = 0 Then
            
            If m_sName = "Trade" Then
                .sCodeID = Format(GetMaxValue("mt_" & m_sName, m_sName & "ID") + 1, "0")
            ElseIf m_sName = "ExWork" Then
                .sCodeID = Format(GetMaxValue("mt_" & m_sName, m_sName & "ID") + 1, "0000")
            ElseIf m_sName = "Color" Then
                .sCodeID = Format(GetMaxValue("mt_" & m_sName, m_sName & "ID"), "00000")
            Else
'                .sCodeID = Format(GetMaxValue("mt_" & m_sName, m_sName & "ID"), "00")
                .sCodeID = Format(GetMinValue("mt_" & m_sName, m_sName & "ID"), "00") '20091215 홍태영 수정
            End If
            
        End If
    End With
    
    g_adoCon.BeginTrans
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_iCode"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, m_sName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, m_sName & "ID")
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, tData.sCodeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, tData.sCode)

        .Execute
    End With
    g_adoCon.CommitTrans
    
    nSql = 1
    
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCode.AddNewcode"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With tData
        sLog(0) = "xp_Code_iCode '" & m_sName & "', '" & m_sName & "ID" & "', '" & .sCodeID & "', '" & .sCode & "'"
    End With
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewCode = False
''         Err.Raise Err.Number, "CCode.AddNewcode", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        AddNewCode = True
    End If

End Function



Public Function AddNewCodeWork(tData As TCode) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%

    On Error GoTo ErrHandler
    g_adoCon.BeginTrans
    
    With tData
        If Len(RTrim(.sCodeID)) = 0 Then
            .sCodeID = Format(GetMaxValue("mt_" & m_sName, m_sName & "ID"), "0000")
            
        End If
    End With
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_iCodeWork"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, tData.sCodeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, tData.sCode)

        .Execute
    End With
    g_adoCon.CommitTrans
    
    nSql = 1
    
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCode.AddNewcodeWork"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With tData
        sLog(0) = "xp_Code_iCode '" & .sCodeID & "', '" & .sCode & "'"
    End With
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewCodeWork = False
''        Err.Raise Err.Number, "CCode.AddNewcodeWork", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        AddNewCodeWork = True
    End If

End Function



Public Function UpdateCode(tData As TCode) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_uCode"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, m_sName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, m_sName & "ID")
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, tData.sCodeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, tData.sCode)

        .Execute
    End With
    
    g_adoCon.CommitTrans
    nSql = 1
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCode.UpdateCode"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With tData
        sLog(0) = "xp_Code_uCode '" & m_sName & "', '" & m_sName & "ID" & "', '" & .sCodeID & "', '" & .sCode & "'"
    End With
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdateCode = False
''        Err.Raise Err.Number, "CCode.UpdateCode", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        UpdateCode = True
    End If
    
End Function



Public Function UpdateCodeWork(tData As TCode) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_uCodeWork"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, tData.sCodeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, tData.sCode)

        .Execute
    End With
    
    g_adoCon.CommitTrans
    nSql = 1
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCode.UpdateCodeWork"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With tData
        sLog(0) = "xp_Code_uCode '" & .sCodeID & "', '" & .sCode & "'"
    End With
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdateCodeWork = False
''        Err.Raise Err.Number, "CCode.UpdateCodeWork", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        UpdateCodeWork = True
    End If
    
End Function


Public Function DeleteCode(sKey As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        If m_sName = "Kind" Or m_sName = "Length" Or m_sName = "Basis" Then
            .CommandText = "xp_Code_dCode"
        Else
            .CommandText = "xp_Code_uDeleteCode"
        End If
    
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, m_sName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, m_sName & "ID")
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sKey)

        .Execute
    End With
    
    g_adoCon.CommitTrans
    nSql = 1
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCode.DeleteCode"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    If m_sName = "Kind" Then
        sLog(0) = "xp_Code_dDeleteCode '" & m_sName & "', '" & m_sName & "ID" & "', '" & sKey & "'"
    Else
        sLog(0) = "xp_Code_uDeleteCode '" & m_sName & "', '" & m_sName & "ID" & "', '" & sKey & "'"
    End If
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeleteCode = False
''        Err.Raise Err.Number, "CCode.DeleteCode", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        DeleteCode = True
    End If
    
End Function

' 2003/12/02 - 유상수
' mt_Hold(보류원인 코드 관련)
Public Function GetHoldCode(sProcessID As String, sHoldClss As String, sHoldID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sHold"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sHoldClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sHoldID)

        Set GetHoldCode = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetHoldCode"), Err.Description
End Function

Public Function AddHoldCode(sProcessID As String, sHoldClss As String, sHoldID As String, _
                            sReason As String, sAbReason As String, sPersonID As String) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_iHold"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sHoldClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sHoldID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, sReason)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sAbReason)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPersonID)
        .Execute
    End With
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    AddHoldCode = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "저장도중 에러"
    AddHoldCode = False
End Function

Public Function GetHoldList(sProcessID As String, sHoldClss As String, sSizeClss As String, _
                            sBig As String, smid As String, ssma As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sHoldByClss"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sHoldClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sSizeClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sBig)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, smid)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ssma)

        Set GetHoldList = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetHoldList"), Err.Description
End Function

Public Function UpdateHoldCode(sProcessID As String, sHoldClss As String, sHoldID As String, _
                            sReason As String, sAbReason As String, sPersonID As String) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_uHold"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sHoldClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sHoldID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, sReason)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sAbReason)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPersonID)
        .Execute
    End With
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    UpdateHoldCode = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "저장도중 에러"
    UpdateHoldCode = False
End Function

Public Function DeleteHoldCode(sProcessID As String, sHoldClss As String, sHoldID As String, sSize As String) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_dHold"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sHoldClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sHoldID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, sSize)
        .Execute
    End With
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    DeleteHoldCode = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "삭제도중 에러"
    DeleteHoldCode = False
End Function

' 공정 관리 SELECT
Public Function GetProcessGroup() As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sProcessGroup"
        Set GetProcessGroup = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

