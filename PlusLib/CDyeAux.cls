VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDyeAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type TDyeAux
    
    DyeAuxID As String     ' 염조제 코드
    DyeAux   As String     ' 염조제 명
    Unit     As String     ' 단위
    UnitCost As Long       ' 단가
    Remark   As String     ' 비고사항
    nKind    As String      ' 구분
End Type

Type TDyeAuxGroup
    GroupID     As String   ' 염조제 그룹 코드
    GroupName   As String   ' 염조제 그룹
    PersonID    As String   ' 작성자
End Type

Type TDyeAuxGroupSub
    GroupID     As String   ' 염조제 그룹 코드
    Seq         As Integer  ' 순위
    DyeAuxID    As String   ' 염조제 코드
End Type

Private Const CLASS_NAME As String = "CDyeAux"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub
'********************************************************
'-- 생지입고 명세서
'********************************************************
Public Function GetDyeAuxSubulsDraft(ByVal sDate As String _
                            , ByVal eDate As String _
                            , ByVal nChkDyeAuxId As Integer _
                            , ByVal sDyeAuxID As String) As ADODB.Recordset
                                
    Dim adoCmd As ADODB.Command, nAffectedRows As Integer
    
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAuxSubul_sDraft"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkDyeAuxId)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sDyeAuxID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamOutput, 1, nAffectedRows)
        
        Set GetDyeAuxSubulsDraft = .Execute
        
        nAffectedRows = .Parameters(4).Value
    End With

    Set adoCmd = Nothing

End Function
Public Function GetDyeAux(sFlag As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_sDyeAux"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sFlag)

        Set GetDyeAux = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDyeAux"), Err.Description
End Function



Public Function GetDyeAuxOne(sDyeAuxID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_sDyeAuxOne"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sDyeAuxID)

        Set GetDyeAuxOne = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDyeAuxOne"), Err.Description
End Function



Public Function AddNewDyeAux(tDA As TDyeAux) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim sNewID$

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_iDyeAux"

        .Parameters.Append .CreateParameter(, adChar, adParamOutput, 4, sNewID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tDA.nKind)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tDA.DyeAuxID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tDA.DyeAux)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tDA.Unit)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, tDA.UnitCost)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tDA.Remark)

        .Execute

        sNewID = .Parameters(0).Value
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDyeAux.AddNewDyeAux"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    With tDA
        sLog(0) = "xp_DyeAux_iDyeAux '" & Left(.DyeAuxID, 1) & "', '" & sNewID & "', '" & .DyeAux & "', '" & _
            .Unit & "', " & .UnitCost & ", '" & .Remark & "'"
    End With

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        AddNewDyeAux = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewDyeAux"), Err.Description
    Else
        Call LogService(sLog)

        AddNewDyeAux = True
    End If
End Function


Public Function UpdateDyeAux(tDA As TDyeAux) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim sNewID$
    
    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_uDyeAux"
        
        sNewID = CStr(tDA.nKind) & tDA.DyeAuxID
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sNewID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tDA.DyeAux)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, tDA.Unit)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, tDA.UnitCost)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tDA.Remark)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDyeAux.UpdateDyeAux"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    With tDA
        sLog(0) = "xp_DyeAux_uDyeAux '" & sNewID & "', '" & .DyeAux & "', '" & .Unit & "', " & .UnitCost & ", " & _
            .Remark & "'"
    End With

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        UpdateDyeAux = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateDyeAux"), Err.Description
    Else
        Call LogService(sLog)

        UpdateDyeAux = True
    End If
End Function

Public Function DeleteDyeAux(sDyeAuxID As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_dDyeAux"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sDyeAuxID)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDyeAux.DeleteDyeAux"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    sLog(0) = "xp_DyeAux_dDyeAux '" & sDyeAuxID & "'"

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        DeleteDyeAux = False
        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteDyeAux"), Err.Description
    Else
        Call LogService(sLog)

        DeleteDyeAux = True
    End If
End Function


Public Function GetDyeAuxGroup() As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_sDyeAuxGroup"

        Set GetDyeAuxGroup = .Execute
        
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDyeAuxGroup"), Err.Description
End Function



Public Function GetDyeAuxGroupSub(sGroupID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_sDyeAuxByGroupID"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sGroupID)

        Set GetDyeAuxGroupSub = .Execute
        
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDyeAuxGroupSub"), Err.Description
End Function


Public Function AddNewDyeAuxGroup(tGroup As TDyeAuxGroup, tGroupSub() As TDyeAuxGroupSub, nCnt As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_iDyeAuxGroup"

        .Parameters.Append .CreateParameter(, adChar, adParamOutput, 4, tGroup.GroupID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tGroup.GroupName)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tGroup.PersonID)
        
        .Execute

        tGroup.GroupID = .Parameters(0).Value
        
        '------------------------------------------------------------------------------------
        .Prepared = True
        
        ' Insert DyeAuxGroup Detail
        .CommandText = "xp_DyeAux_iDyeAuxGroupSub"

        For i = 0 To nCnt - 1
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tGroup.GroupID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tGroupSub(i).Seq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tGroupSub(i).DyeAuxID)

            .Execute
          
        Next i
    End With
    
    AddNewDyeAuxGroup = True
    
    Exit Function

ErrHandler:
    bError = True

    AddNewDyeAuxGroup = False
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDyeAux.AddNewDyeAuxGroup"
End Function





Public Function UpdateDyeAuxGroup(tGroup As TDyeAuxGroup, tGroupSub() As TDyeAuxGroupSub, nCnt As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_uDyeAuxGroup"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tGroup.GroupID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, tGroup.GroupName)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tGroup.PersonID)
        
        .Execute
        
        '------------------------------------------------------------------------------------
        .Prepared = True
        
        .CommandText = "xp_DyeAux_dDyeAuxGroupSub"
        
        Call ClearParameter(adoCmd)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tGroup.GroupID)
        
        .Execute
        
        '------------------------------------------------------------------------------------

        ' Insert DyeAuxGroup Detail
        .CommandText = "xp_DyeAux_iDyeAuxGroupSub"

        For i = 0 To nCnt - 1
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tGroup.GroupID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tGroupSub(i).Seq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tGroupSub(i).DyeAuxID)

            .Execute
            
        Next i

    End With
    UpdateDyeAuxGroup = True
    Exit Function
    
ErrHandler:
    bError = True
    UpdateDyeAuxGroup = False
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CDyeAux.UpdateDyeAuxGroup"
End Function



Public Function DeleteDyeAuxGroup(sGroupID As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_DyeAux_dDyeAuxGroup"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sGroupID)
        
        .Execute

    End With
    
    DeleteDyeAuxGroup = True
    Exit Function
    

ErrHandler:
    bError = True

    DeleteDyeAuxGroup = False
End Function


