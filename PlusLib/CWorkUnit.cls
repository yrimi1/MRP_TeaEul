VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWorkUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type TWorkUnit
    sCardID       As String
    sSplitID      As String
    sWorkUnitID   As String
    nWorkUnitSeq  As Integer
    sBatJaNo    As String
End Type

Private Const CLASS_NAME As String = "CWorkUnit"

Public Property Let Connection(g_adoConnect As Object)
    Set g_adoCon = g_adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Public Function GetProcess(Optional nChkFlag As Integer = 0) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Common_sProcess"
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 2, nChkFlag)
        
        Set GetProcess = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetOrder(nChkCustomID As Integer, sCustomID As String, _
                    nChkArticleID As Integer, sArticleID As String, _
                    nChkOrder As Integer, sOrder As String, _
                    nChkCardID As Integer, sCardID As String, _
                    nChkProcessID As Integer, sProcessID As String) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WorkUnit_sOrder"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCardID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkProcessID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sProcessID)

        Set GetOrder = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetBatJaNo(sBatJaNo As String) As Boolean
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_WorkUnit_sBatJaNo"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sBatJaNo)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, 0)
        
        .Execute
        
        GetBatJaNo = IIf(.Parameters(1).Value = 0, False, True)
        
    End With
    
    Set adoCmd = Nothing
End Function

Public Function ModifyWorkUnit(tWork() As TWorkUnit) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%
    Dim sWorkUnitID$
    
    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    
    bError = False
    iSQL = 0
    
    g_adoCon.BeginTrans
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        For i = 0 To UBound(tWork)
            .CommandText = "xp_WorkUnit_uWorkUnit"
            .Prepared = True
            
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, tWork(i).sCardID)
            .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, tWork(i).sSplitID)
            .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, tWork(i).sBatJaNo)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInputOutput, 10, sWorkUnitID)
            
            .Execute
            sWorkUnitID = .Parameters(3).Value
            
            Call ClearParameter(adoCmd)
            iSQL = iSQL + 1
        Next i
    End With
    g_adoCon.CommitTrans
    
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CWorkUnit.ModifyWorkUnit"
LogMessage:
    Set adoCmd = Nothing
    ReDim sLog(UBound(tWork))

    For i = 0 To UBound(tWork)
        With tWork(i)
            sLog(i) = "xp_WorkUnit_iWorkUnit '" & .sCardID & "', '" & .sSplitID & "', '" & .sBatJaNo & "', '" & sWorkUnitID & "' "
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        ModifyWorkUnit = False
''        Err.Raise Err.Number, CLASS_NAME & ".ModifyWorkUnit", Err.Description
    Else
        Call LogService(sLog)

        ModifyWorkUnit = True
    End If
End Function
