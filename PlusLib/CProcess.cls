VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type TProcess
    sProcessID      As String '[1] 공정 코드
    sProcess        As String '[2] 공정 명
End Type

Type TMachine
    sProcessID      As String '[1] 공정 코드
    sMachineID      As String '[2] 기계 코드
    sMachine        As String '[3] 기계 명
    sMachineNO      As String '[4] 기계 호기
End Type


Type TProcessResult
    sCardID         As String
    sSplitID        As String
    nWorkSeq        As Integer
    sProcessID      As String
    nTemper         As Integer
    nVelocity       As Integer
    nOverFeed       As Single
    nDensity        As Integer
    sSetting        As String
    sWorkCon        As String
    sCodeID         As String
    sDryID          As String
    sSettingClss    As String
    sSideClss       As String
    nWind           As Integer
    nGas            As Integer
    nRPM            As Integer
    sDyeAuxID       As String
    sRefineClss     As String
    nPepaBon1       As Integer
    nPepaBon2       As Integer
    nPepaBon3       As Integer
    nPepaBon4       As Integer
    nTension        As Integer
    nPressure1      As Integer
    nPressure2      As Integer
    nPressure3      As Integer
    nRealLoss       As Single
End Type


Public Enum EPROCESSCODE
    ' wk_Moso
    PC_Moso = "1901"
    
    ' wk_CPBPre
    PC_Pre = "2201"
    PC_1stHobal = "2204"        ' 호발
    PC_Pufiry = "2207"          ' 정련
    PC_1stPurify = "2210"       ' 1차정련
    PC_2ndPurify = "2213"       ' 2차정련
    PC_2ndReduce = "2217"       ' 2차감량
    PC_LBoxPre = "2220"         ' LBox 전처리
    PC_CPBPre = "2223"          ' CPB 전처리
    PC_NewSTPre = "2224"        ' 신 ST 전처리
'
    ' wk_Refine
    PC_REFINE = "2501"          ' 수세
    
    ' wk_Mercerizing
    PC_SK = "2801"              '
    PC_NewST = "2804"           ' 신 ST
    PC_OBoxSK = "2807"          '
    
    ' wk_Tenter
    PC_Setting = "3101"         '셋팅
    PC_WidthLine = "3104"        '폭줄
    PC_FinalSetting = "7601"    '가공
    
    ' wk_Peach
    PC_Peach = "3401"           'Peach
    
    ' CPB 염색
    PC_CPB = "4001"
    
    ' wk_Dry
    PC_Dry = "7605"             '건조
    PC_WRSetting = "7603"    'W/R가공
    
    ' wk_Shampu
    PC_Shampu = "6701"          '샴푸
    
    ' 시작과 끝 시간만 올림
    PC_Airo = "7001"            'Airo
    PC_Calender = "7301"        '카렌다
    
    PC_Sanfor = "7901"

End Enum


' 설비 코드
'Public Enum EPlantCODE
'    ' 정련 CPB - 1차정련, 2차정련, 2차감량, 효소호발
'    MC_CPBPre = "2204"
'    ' CPB 염색
'    MC_CPBDye = "4001"
'    ' 수세기 - 수세, S/K
'    MC_REFINE = "2501"
'    ' PEACH
'    MC_PEACH = "3401"
'    ' DRY
'    MC_DRY = "6401"
'    ' 텐터
'    MC_TENTER = "3101"
'
'End Enum


Private Const CLASS_NAME As String = "CProcess"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub
' 공정 관리 SELECT
Public Function GetProcess() As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sProcess"
        Set GetProcess = .Execute
    End With
    
    Set adoCmd = Nothing
End Function


' 설비명 받아오기
Public Function GetPlant() As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Work_sMachine"
        

        Set GetPlant = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

End Function



' 공정 관리 SELECT
Public Function GetWorkProcess() As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Work_sProcess"
        
        Set GetWorkProcess = .Execute
    End With
    
    Set adoCmd = Nothing
End Function



' 기계 관리 SELECT
Public Function GetMachine(Optional NewKey As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Process_sMachine"
        If Len(NewKey) > 0 Then
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, NewKey)
        End If
        Set GetMachine = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

' 설비별 호기
Public Function GetMachineByPlant(sMachine As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Work_sMachineNO"
        
        ' 설비 이름으로 검색
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sMachine)
        
        Set GetMachineByPlant = .Execute
                
    End With
    Set adoCmd = Nothing

End Function


Public Function AddNewProcess(NewItem As TProcess) As Boolean
    Dim SQL As String
    
    With NewItem
        If Len(.sProcessID) <> 2 Then
            .sProcessID = Format(GetMaxValue("[mt_Process]", "[ProcessID]"), "00")
        End If
            
        SQL = "INSERT INTO [mt_Process] (ProcessID, Process) " & _
              " VALUES (" & CheckNull(.sProcessID) & ", " & CheckNull(.sProcess) & ") "
    End With

    AddNewProcess = HandleDB(SQL)
End Function

Public Function UpdateProcess(NewItem As TProcess) As Boolean
    Dim SQL As String

    With NewItem
        SQL = "UPDATE [mt_Process] " & _
              "SET Process = " & CheckNull(.sProcess) & " " & _
              "WHERE ProcessID = " & CheckNull(.sProcessID)
    End With

    UpdateProcess = HandleDB(SQL)
End Function

Public Function DeleteProcess(NewKey As String) As Boolean
    Dim SQL As String

    SQL = "DELETE FROM [mt_Process] WHERE ProcessID = '" & NewKey & "' "

    DeleteProcess = HandleDB(SQL)
End Function

Public Function AddNewMachine(NewItem As TMachine) As Boolean
    Dim SQL As String
    
    With NewItem
        If Len(.sMachineID) <> 3 Then
            .sMachineID = Format(GetMaxValue("[mt_Machine]", "[MachineID]", "WHERE ProcessID = " & CheckNull(.sProcessID) & " "), "00")
        End If
            
        SQL = "INSERT INTO [mt_Machine] (ProcessID, MachineID, Machine, MachineNO) " & _
              " VALUES (" & CheckNull(.sProcessID) & ", " & CheckNull(.sMachineID) & ", " & _
              CheckNull(.sMachine) & ", " & CheckNull(.sMachineNO) & " ) "
    End With
    AddNewMachine = HandleDB(SQL)
End Function

Public Function UpdateMachine(NewItem As TMachine) As Boolean
    Dim SQL As String

    With NewItem
        SQL = "UPDATE [mt_Machine] " & _
              "SET Machine = " & CheckNull(.sMachine) & ", " & _
              "MachineNO = " & CheckNull(.sMachineNO) & " " & _
              "WHERE ProcessID = " & CheckNull(.sProcessID) & " AND MachineID = " & CheckNull(.sMachineID)
    End With

    UpdateMachine = HandleDB(SQL)
End Function

Public Function DeleteMachine(NewKey1 As String, NewKey2 As String) As Boolean
    Dim SQL As String

    SQL = "DELETE FROM [mt_Machine] WHERE ProcessID = '" & NewKey1 & "' AND MachineID = '" & NewKey2 & "' "

    DeleteMachine = HandleDB(SQL)
End Function

Public Function GetProcessByOrder(nChk1 As Integer, sFDate As String, sTDate As String, nChk2 As Integer, _
sCustomID As String, nChk3 As Integer, sOrderID As String, nChk4 As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    Dim nChkOrder As Integer
    
    If nChk3 = 1 Then
        If nChk4 = 1 Then
            nChkOrder = 1
        ElseIf nChk4 = 2 Then
            nChkOrder = 2
        End If
    Else
        nChkOrder = 0
    End If
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Process_sProcessByOrder"
        
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChk1)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sFDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sTDate)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChk2)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 15, sOrderID)
        Set GetProcessByOrder = .Execute
    End With

    Set adoCmd = Nothing
End Function



'---------------------------------------------------------
' 공정관리 일지
'---------------------------------------------------------
Public Function GetProcessResultMgr(ByVal nChkDate As Integer, ByVal sDate As String, ByVal eDate As String _
                                , ByVal nChkOrder As Integer, ByVal sOrderID As String _
                                , ByVal nChkCustom As Integer, ByVal sCustomID As String _
                                , ByVal nChkArticle As Integer, ByVal sArticleID As String _
                                , ByVal nChkCard As Integer, ByVal sFromCardID As String, ByVal sToCardID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_ProcessResultMgr_sDraft"
        
        .Parameters.Append .CreateParameter(.CommandText, adTinyInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(.CommandText, adTinyInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adTinyInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adTinyInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adTinyInt, adParamInput, 1, nChkCard)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 12, sFromCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 12, sToCardID)
        Set GetProcessResultMgr = .Execute
    End With

    Set adoCmd = Nothing

End Function
Public Function GetResultByProcess(sSDate As String, sEDate As String, sProcessID As EPROCESSCODE, nChkMachineID As Integer, _
    sMachineID As String, nChkTeamID As Integer, sTeamID As String, nChkOrder As Integer, sOrder As String, _
    nChkCustomID As Integer, sCustomID As String, nChkArticleID As Integer, sArticleID As String) As Recordset
    
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    ' 텐터, 건조, 수세, 모소, m/c, cpb전처리,  peach, 샴푸
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        Select Case sProcessID
            Case PC_Moso
                .CommandText = "xp_Work_sResultWithMoso"
                
            ' wk_CPBPre
            Case PC_Pre     ' 전처리
                .CommandText = "xp_Work_sResultWithCPBPre_Pre"
                
            Case PC_1stHobal        '1차 호발
                .CommandText = "xp_Work_sResultWithCPBPre_1stHobal"

            Case PC_Pufiry          '정련
                .CommandText = "xp_Work_sResultWithCPBPre_Purify"  'OK

            Case PC_1stPurify       '1차정련
                .CommandText = "xp_Work_sResultWithCPBPre_1stPurify"

            Case PC_2ndPurify       ' 2차정련
                .CommandText = "xp_Work_sResultWithCPBPre_2ndPurify"
                
            Case PC_2ndReduce       '2차감량
                .CommandText = "xp_Work_sResultWithCPBPre_2ndReduce"
                
            ' wk_Refine
            Case PC_REFINE          ' 수세
                .CommandText = "xp_Work_sResultWithRefine"
                
            ' wk_Mercerizing
            Case PC_SK
                .CommandText = "xp_Work_sResultWithMercerizing"
                
            ' wk_Tenter
            Case PC_Setting         '셋팅
                .CommandText = "xp_Work_sResultWithTenter_Setting"
                
            Case PC_WidthLine       '폭줄
                .CommandText = "xp_Work_sResultWithTenter_WidthLine"
                
            Case PC_FinalSetting    '가공
                .CommandText = "xp_Work_sResultWithTenter_Final"
                
            ' wk_Peach
            Case PC_Peach           'Peach
                .CommandText = "xp_Work_sResultWithPeach"
                
            ' wk_Dry
            Case PC_Dry             '건조
                .CommandText = "xp_Work_sResultWithDry"
                
            Case PC_WRSetting             'W/R가공
                .CommandText = "xp_Work_sResultWithTenter_WR"
                
            ' wk_Shampu
            Case PC_Shampu          '샴푸
                .CommandText = "xp_Work_sResultWithShampu"
                
            Case PC_CPB
                .CommandText = "xp_Work_sResultWithCPBDye"
            
            ' 시작과 끝 시간만 올림
            Case PC_Airo            'Airo
                .CommandText = "xp_Work_sResultWithAiro"

            Case PC_Calender        '카렌다
                .CommandText = "xp_Work_sResultWithCalender"

        End Select

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkMachineID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 3, Format(sMachineID, "00"))
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


' 설비별 실적
Public Function GetResultByPlant(ByVal sSDate As String, ByVal sEDate As String, ByVal sProcessID As EPROCESSCODE _
                                , ByVal nChkMachineID As Integer, ByVal sMachineID As String _
                                , ByVal nChkTeamID As Integer, ByVal sTeamID As String, ByVal nChkOrderID As Integer, ByVal sOrderID As String _
                                , ByVal nChkCustomID As Integer, ByVal sCustomID As String _
                                , ByVal nChkArticleID As Integer, ByVal sArticleID As String) As ADODB.Recordset

    Dim adoCmd As ADODB.Command

    ' 텐터 , 건조, 수세, 모소, m / c, cpb전처리, peach, 샴푸
     
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        Select Case sProcessID
            Case PC_1stHobal
                .CommandText = "xp_Work_sResultWithMachineCPBPre"
            Case PC_REFINE
                .CommandText = "xp_Work_sResultWithMachineRefine"
            Case PC_Setting
                .CommandText = "xp_Work_sResultWithMachineTenter"
            Case PC_Peach
                .CommandText = "xp_Work_sResultWithMachinePEACH"
            Case PC_CPB
                .CommandText = "xp_Work_sResultWithMachineCPBDye"
            Case PC_Dry
                .CommandText = "xp_Work_sResultWithMachineDRY"
            Case PC_Pufiry          '정련
                .CommandText = "xp_Work_sResultWithCPBPre_Purify"
        End Select

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, Format(sProcessID, "0000"))
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkMachineID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 3, Format(sMachineID, "00"))
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkTeamID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sTeamID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        
        Set GetResultByPlant = .Execute
    End With
    Set adoCmd = Nothing

End Function


Public Function UpdateProcessResult(NewResult As TProcessResult) As Boolean
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
        .CommandText = "xp_Work_uResultByProcess"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewResult.sCardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, IIf(Len(NewResult.sSplitID) = 0, " ", NewResult.sSplitID))
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 8, NewResult.nWorkSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewResult.sProcessID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nTemper)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nVelocity)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, NewResult.nOverFeed)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nDensity)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewResult.sWorkCon)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewResult.sCodeID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewResult.sDryID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, NewResult.sSettingClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, NewResult.sSideClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nWind)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nGas)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nRPM)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewResult.sDyeAuxID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 6, NewResult.sRefineClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nPepaBon1)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nPepaBon2)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nPepaBon3)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nPepaBon4)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nTension)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nPressure1)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nPressure2)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewResult.nPressure3)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, NewResult.nRealLoss)
             
        
        .Execute
    
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CProcess.UpdateProcessResult"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)
    With NewResult
        sLog(0) = "xp_Work_uResultByProcess '" & .sCardID & "', '" & .sSplitID & "', " & .nWorkSeq & ", '" & .sProcessID & "', " & _
                        .nTemper & ", " & .nVelocity & ", " & .nOverFeed & ", " & .nDensity & ", '" & .sSetting & "', '" & _
                        .sWorkCon & "', '" & .sCodeID & "', '" & .sDryID & "', '" & .sSettingClss & "', " & .nWind & ", " & _
                        .nGas & ", " & .nRPM & ", '" & .sDyeAuxID & "', '" & .sRefineClss & "', " & .nPepaBon1 & ", " & _
                        .nPepaBon2 & ", " & .nPepaBon3 & ", " & .nTension & ", " & .nPressure1 & ", " & .nPressure2 & ", " & _
                        .nPressure3 & ", " & .nRealLoss
        
    End With
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdateProcessResult = False
''        Err.Raise Err.Number, "CProcess.UpdateProcessResult", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        UpdateProcessResult = True
    End If

End Function



' 카드별 실적
Public Function GetResultByCard(sCardID As String, sSplitID As String) As ADODB.Recordset

    Dim adoCmd As ADODB.Command
     
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        .CommandText = "xp_Work_sResultByCard"
                
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)
        
        Set GetResultByCard = .Execute
    End With
    Set adoCmd = Nothing

End Function


Public Function GetWaitCardList(nChkOrder As Integer, sOrder As String, nChkCustomID As Integer, sCustomID As String, _
                                nChkArticleID As Integer, sArticleID As String, nChkProcess As Integer, sProcessID As String) As ADODB.Recordset

    Dim adoCmd As ADODB.Command
     
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        If nChkProcess = 0 Then
            .CommandText = "xp_Work_sWaitCardList"
            
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrder)
            .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, sOrder)
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticleID)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        Else
            .CommandText = "xp_Work_sWaitCardListByProcess"
            
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrder)
            .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, sOrder)
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
            .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticleID)
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
            .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sProcessID)
        End If
        
        Set GetWaitCardList = .Execute
    End With
    Set adoCmd = Nothing

End Function
