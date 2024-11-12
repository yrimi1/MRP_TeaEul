VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDefect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TDefect
    DefectID       As String   ' [0] 불량 코드
    Display1       As String   ' [1] 단말기 Display(1)
    Display2       As String   ' [2] 단말기 Display(2)
    Display3       As String   ' [3] 단말기 Display(3)
    KDefect        As String   ' [4] 불량명(한글)
    EDefect        As String   ' [5] 불량명(영문)
    TagName        As String   ' [6] Tag Name
    KindID         As String   ' [7] 불량 종류
    KindIDSub      As String   ' [7] 불량 종류 세부
End Type

Public Type TDefectSub
    DefectID       As String   ' [0] 불량코드
    ButtonSeq      As Integer  ' [1] 버튼 위치
    Demerit        As Integer  ' [2] 벌점
    Loss           As Integer  ' [3] 보상
End Type

Private Const CLASS_NAME As String = "CDefect"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetDefect(sBasisID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sDefect"

        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 1, sBasisID)

        Set GetDefect = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDefect"), Err.Description
End Function

Public Function GetDefectByLang(nChkLang As Integer, sBasisID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sDefectByLang"

        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkLang)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 1, sBasisID)

        Set GetDefectByLang = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDefectByLang"), Err.Description
End Function
Public Function AddNewDefect(tDef As TDefect) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Defect_iDefect"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Left(tDef.DefectID, 1))
        .Parameters.Append .CreateParameter(, adChar, adParamOutput, 3, tDef.DefectID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, tDef.Display1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, tDef.Display2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, tDef.Display3)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tDef.KDefect)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tDef.EDefect)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tDef.TagName)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tDef.KindID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tDef.KindIDSub)

        .Execute

        tDef.DefectID = .Parameters(1).Value
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDefect.AddNewDefect"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    With tDef
        sLog(0) = "xp_Defect_iDefect '" & Left(.DefectID, 1) & "', '" & .DefectID & "', '" & .Display1 & "', '" & _
            .Display2 & "', '" & .Display3 & "', '" & .KDefect & "', '" & .EDefect & "', '" & .TagName & "', '" & .KindID & "', '" & .KindIDSub & "' "
    End With

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        AddNewDefect = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewDefect"), Err.Description
    Else
        Call LogService(sLog)

        AddNewDefect = True
    End If
End Function



Public Function AddNewDefectSub(NewDefectSub() As TDefectSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%
    
    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    
    bError = False
    iSQL = 0
    
    g_adoCon.BeginTrans
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Defect_dDefectSub"
        .Prepared = True
        
        .Execute
        
        iSQL = iSQL + 1
        
        .CommandText = "xp_Defect_iDefectSub"
        For i = 0 To UBound(NewDefectSub)
            Call ClearParameter(adoCmd)
            
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewDefectSub(i).DefectID)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewDefectSub(i).ButtonSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewDefectSub(i).Demerit)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewDefectSub(i).Loss)
    
            .Execute
            
        Next i
    End With
    g_adoCon.CommitTrans
    GoTo LogMessage

ErrHandler:
    bError = True
    g_adoCon.RollbackTrans
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDefect.AddNewDefectSub"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(UBound(NewDefectSub) + 1)
    
    sLog(0) = "xp_Defect_dDefectsub"

    For i = 1 To UBound(NewDefectSub)
        With NewDefectSub(i - 1)
            sLog(i) = "xp_Defect_iDefectSub '" & .DefectID & "', " & .ButtonSeq & ", " & .Demerit & ", " & .Loss
        End With
    Next i
    

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        AddNewDefectSub = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewDefectSub"), Err.Description
    Else
        Call LogService(sLog)

        AddNewDefectSub = True
    End If
End Function


Public Function UpdateDefect(tDef As TDefect) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Defect_uDefect"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tDef.DefectID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, tDef.Display1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, tDef.Display2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, tDef.Display3)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tDef.KDefect)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tDef.EDefect)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tDef.TagName)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tDef.KindID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tDef.KindIDSub)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDefect.UpdateDefect"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    With tDef
        sLog(0) = "xp_Defect_uDefect '" & .DefectID & "', '" & .Display1 & "', '" & .Display2 & "', '" & .Display3 & "', '" & _
            .KDefect & "', '" & .EDefect & "', '" & .TagName & "', '" & .KindID & "', '" & .KindIDSub & "' "
    End With

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        UpdateDefect = False
        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateDefect"), Err.Description
    Else
        Call LogService(sLog)

        UpdateDefect = True
    End If
End Function

Public Function UpdateDefectSub(NewDefectSub() As TDefectSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim iLoop%
    
    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Defect_uDefectSub"
        .Prepared = True
        
        For iLoop = 0 To UBound(NewDefectSub)
            Call ClearParameter(adoCmd)
            
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewDefectSub(iLoop).DefectID)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewDefectSub(iLoop).ButtonSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewDefectSub(iLoop).Demerit)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewDefectSub(iLoop).Loss)
    
            .Execute
            
        Next iLoop
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDefect.UpdateDefectSub"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(UBound(NewDefectSub))
    

    For iLoop = 0 To UBound(NewDefectSub)
        With NewDefectSub(iLoop)
            sLog(iLoop) = "xp_Defect_uDefectSub '" & .DefectID & "', " & .ButtonSeq & ", " & .Demerit & ", " & .Loss
        End With
    Next iLoop
    

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        UpdateDefectSub = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateDefectSub"), Err.Description
    Else
        Call LogService(sLog)

        UpdateDefectSub = True
    End If
End Function

Public Function DeleteDefect(sDefectID As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Defect_dDefect"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sDefectID)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDefect.DeleteDefect"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
 
    sLog(0) = "xp_Defect_dDefect '" & sDefectID & "'"

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        DeleteDefect = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteDefect"), Err.Description
    Else
        Call LogService(sLog)

        DeleteDefect = True
    End If
End Function
