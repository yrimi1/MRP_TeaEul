VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRapid_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' 염색일지
Type TWkRapid
    wkResultDT         As String
    wkProcID           As String
    wkMachID           As String
    wkRapidSeq         As Integer
    WorkClss           As String
    RapidClss          As String
    PatternID          As String
    InMethod           As String
    UnitWght           As Integer
    WkRoll             As Integer
    WkQty              As Integer
    TeamID             As String
    PersonID           As String
    StartDate          As String
    StartTime          As String
    EndDate            As String
    EndTime            As String
    DyeSchID           As Integer
    DyeSeq             As Integer
    Remark             As String
    HoldReason         As String
End Type

' 염색일지 Sub Type
Type TWkRapidSUB
    wkResultDT         As String
    wkProcID           As String
    wkMachID           As String
    wkRapidSeq         As Integer
    CardID             As String
    SplitID            As String
    ReWorkClss         As String
    ReWorkID           As String
    PlanSeq             As Integer
End Type



Private Const CLASS_NAME As String = "CRapid_NEW"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub



''xp_Rapid_iRapid
''         wkResultDT         as string
''         wkProcID           as string
''         wkMachID           as string
''         wkRapidSeq         as integer
''         WorkClss           as string
''         RapidClss          as string
''         PatternID          as string
''         InMethod           as string
''         UnitWght           as integer
''         WkRoll             as integer
''         WkQty              as integer
''         TeamID             as string
''         PersonID           as string
''         StartDate          as string
''         StartTime          as string
''         EndDate            as string
''         EndTime            as string
''         DyeSchID           as integer
''         DyeSeq             as integer
''         Remark             as string
''         HoldReason         as string

Public Function InsertRapid(ByRef TWkRapid As TWkRapid _
                         , ByRef TWkRapidSUB() As TWkRapidSUB _
                         , ByVal nCount As Integer _
                         , ByVal JobMode As String) As Boolean
    Dim iChkDye As Integer, II As Integer
    Dim nSeq As Integer
    Dim adoCmd As ADODB.Command, sUseClss As String, AA As Integer
    Dim sWorkUnitID As String, nWorkUnitSeq As Integer
    On Error GoTo ErrHandler

''    If sClss = "염색" Then
''        iChkDye = 0
''    Else
''        iChkDye = 1
''    End If

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_iRapid"
    
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TWkRapid.wkResultDT)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TWkRapid.wkProcID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TWkRapid.wkMachID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInputOutput, 2, TWkRapid.wkRapidSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, TWkRapid.WorkClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, TWkRapid.RapidClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, TWkRapid.PatternID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 6, TWkRapid.InMethod)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TWkRapid.UnitWght)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TWkRapid.WkRoll)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, TWkRapid.WkQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TWkRapid.TeamID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TWkRapid.PersonID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TWkRapid.StartDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TWkRapid.StartTime)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TWkRapid.EndDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TWkRapid.EndTime)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, TWkRapid.DyeSchID)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, TWkRapid.DyeSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, TWkRapid.Remark)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TWkRapid.HoldReason)
        
        .Execute
        nSeq = .Parameters(3).Value
        
        If TWkRapid.EndDate = "" Then
            sUseClss = "작업"
        Else
            sUseClss = "대기"
        End If
        
        If nCount > 0 Then
            .CommandText = "xp_Rapid_iRapidSub"
            .Prepared = True
    
            For II = 0 To UBound(TWkRapidSUB) - 1
                Call ClearParameter(adoCmd)
    
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TWkRapidSUB(II).wkResultDT)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TWkRapidSUB(II).wkProcID)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TWkRapidSUB(II).wkMachID)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nSeq)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, TWkRapidSUB(II).CardID)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, TWkRapidSUB(II).SplitID)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sUseClss)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, JobMode)
                .Parameters.Append .CreateParameter(, adChar, adParamInputOutput, 10, sWorkUnitID)
                .Parameters.Append .CreateParameter(, adSmallInt, adParamInputOutput, 2, nWorkUnitSeq)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TWkRapidSUB(II).ReWorkClss)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TWkRapidSUB(II).ReWorkID)
                
                .Execute AA
                
                sWorkUnitID = .Parameters(8).Value
                nWorkUnitSeq = .Parameters(9).Value + 1
            Next II
        End If
        
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    InsertRapid = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "저장도중 에러"
    InsertRapid = False
End Function



' wkRapid와 관련된 Card List가져오기
Public Function GetRapidCardList(ByVal ResultDate As String, ByVal ProcID As String, ByVal MachID As String, ByVal RapidSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_wkRapid_CardList"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, ResultDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, ProcID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, MachID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, RapidSeq)
        
        Set GetRapidCardList = .Execute
    End With

    Set adoCmd = Nothing

End Function

' wiRapid 의 DyeSchID에 따른 작업 조건 가져오기
Public Function GetWiRapidSch(ByVal nDyeSchID As Integer, ByVal nDyeSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_swiRapidSch"

        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nDyeSchID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nDyeSeq)
        
        Set GetWiRapidSch = .Execute
    End With

    Set adoCmd = Nothing

End Function
' 염색대기에 있는 카드 내역 리스트
Public Function GetRapidScheduling_END() As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyeScheduleList_END"

        
        Set GetRapidScheduling_END = .Execute
    End With

    Set adoCmd = Nothing
End Function


' -- 작업중인 것, 작업완료된 카드 작업 취소
Public Function SetRapiddCancel(ByVal sResltDT As String, ByVal sProcID As String, ByVal sMachID As String, ByVal nRapidSeq As Integer) As Boolean

' xp_wkRapid_dCancel

    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command
    SetRapiddCancel = True
    On Error GoTo ErrHandle
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_wkRapid_dCancel"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sResltDT)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sProcID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sMachID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nRapidSeq)
        
        .Execute
    End With

    Set adoCmd = Nothing
    Exit Function
    
ErrHandle:
    SetRapiddCancel = False
End Function
' 염색대기에 있는 카드 내역 리스트
Public Function GetRapidScheduling_NOW(nChkValue As Integer) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyeScheduleList_NOW"

        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkValue)
        
        Set GetRapidScheduling_NOW = .Execute
    End With

    Set adoCmd = Nothing
End Function

' 염색작업중인 카드 내역 리스트
' 0: 작업중, 1:완료
Public Function GetRapidScheduling_ING(ByVal sCheck As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyeScheduleList_ING"
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, sCheck)
        Set GetRapidScheduling_ING = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetRapidScheduling_ReWork(sCardID As String, sSplitID As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyeScheduleList_ReWork"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)
        
        Set GetRapidScheduling_ReWork = .Execute
    End With

    Set adoCmd = Nothing
End Function

