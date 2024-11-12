VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************************************************
'** System 명 : PlusLib2
'** 모듈명    : 사원관리 DB처리용
'** Author    : Wizard
'** 작성자    :
'** 내용      :
'** 생성일자  :
'------------------------------------------------------------------------------------------------------------------
' * 변경이력
'------------------------------------------------------------------------------------------------------------------
' 일자        작업자  요청자          요청번호           요청내용 및 변경내용
'------------------------------------------------------------------------------------------------------------------
' 2013.12.12  오승욱                 S_201312_삼우_99    지번주소에서 도로명 주소로 입력가능하게-사원 관련 Type변수 수정,AddNewPerson  함수,UpdatePerson 함수 수정
'*******************************************************************************

Option Explicit

Type TPerson
    sPersonID       As String    ' [0] 코드 (4)
    sUserID         As String    ' [1] Username
    sPassword       As String    ' [2] Password
    sName           As String    ' [3] 성명
    sDepartID       As String    ' [5] 부서
    sDutyID         As String    ' [6] 직책 (0:사장,1:전무,2:상무,3:이사,4:부장,5:차장,6:과장,7:대리,8:계장,9:관리직사원,10:생산직 사원)
    sTeamID         As String    ' 작업조
    sStartDate      As String    ' [7] 입사일자
    sEndDate        As String    ' [8] 퇴사일자
    sRegistID       As String    ' [9] 주민등록번호
    sHandPhone      As String    ' [10] 핸드폰
    sPhone          As String    ' [11] 전화번호
    sBirthDay       As String    ' [12] 생년월일
    sSolarClss      As String    ' [13] 양/음력 구분
    sZipCode        As String    ' [14] 우편번호
    'S_201312_삼우_99 에 의한 추가---------------------------------------------------
    sOldNNewClss        As String '[15]  도로명,지번주소 구분 0:도로명, 1:지번
    sGunMoolMngNo        As String '[16]  건물관리 고유식별번호
    sAddress1       As String '[17] 도로명주소1
    sAddress2       As String '[18] 도로명주소2
    sAddressAssist       As String '[19] 도로명보조주소
    '--------------------------------------------------------------------------------
    'S_201312_삼우_99 에 의한 수정(OLD:sAddress1)
    sAddressJiBun1       As String    ' [20] 집 주소1
    'S_201312_삼우_99 에 의한 수정(OLD:sAddress1)
    sAddressJiBun2       As String    ' [21] 집ㅉ 주소2
    sEMail          As String        ' [22] 이메일
    sRemark         As String    ' [23] 비고사항
    sEName          As String    ' [24] 영문이름
    sSMSYN          As String   '  [25] SMS문자 전송여부(Y:문자 전송 대상, N:대상 아님)
End Type


Type TPersonMachine
    sPersonID   As String
    sProcessID  As String
    sMachineID  As String
    
End Type

Private Const CLASS_NAME As String = "CMenu"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetPerson(Optional sDepartID As String = "") As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    Dim nChkDepartID%

    On Error GoTo ErrHandler

    nChkDepartID = IIf(Len(sDepartID) = 0, 0, 1)

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_sPerson"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkDepartID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sDepartID)

        Set GetPerson = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set GetPerson = Nothing
    Err.Raise Err.Number, "CPerson.GetPerson", Err.Description
End Function


Public Function GetPersonMachine(sPersonID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_sPersonMachine"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sPersonID)
        
        Set GetPersonMachine = .Execute
    End With

    Set adoCmd = Nothing

End Function


Public Function AddNewPerson(stData As TPerson, stMenu() As TUSERMENU, nCnt As Integer, _
                                stMachine() As TPersonMachine, nMachineCnt As Integer) As Boolean
    Dim PersonID As String
    Dim iLoop%, nSql%
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_iPerson"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInputOutput, 15, stData.sPersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, stData.sUserID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sPassword)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, stData.sName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, stData.sDepartID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, stData.sDutyID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sStartDate)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sEndDate)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 13, stData.sRegistID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 13, stData.sHandPhone)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, stData.sPhone)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sBirthDay)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, stData.sSolarClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 7, stData.sZipCode)
        'S_201312_삼우_99 에 의한 추가-------------------------------------------------------
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, stData.sOldNNewClss)       '도로명,지번주소 구분 0:도로명, 1:지번
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, stData.sGunMoolMngNo)     '건물관리 고유식별번호
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddress1)      '도로명 주소1
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddress2)      '도로명 주소2
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddressAssist)    '도로명 보조 주소
        '----------------------------------------------------------------------------------------
        'S_201312_삼우_99 에 의한 수정(OLD:stData.sAddress1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sAddressJiBun1)
        'S_201312_삼우_99 에 의한 수정(OLD:stData.sAddress2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, stData.sAddressJiBun2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sEMail)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sRemark)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, stData.sTeamID)               '작업조
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, stData.sEName)           '영문이름
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, stData.sSMSYN)                '문자전송 대상여부
        
        .Execute
        
        PersonID = .Parameters(0).Value
    End With
    
    nSql = 1
    
    If nCnt <> 0 Then
        With adoCmd     ' 사용자별 메뉴 입력
            .CommandText = "xp_Menu_iUserMenu"
            .Prepared = True
        
            For iLoop = 0 To nCnt - 1
                Call ClearParameter(adoCmd)
        
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, PersonID)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, stMenu(iLoop).sMenuID)
                .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, stMenu(iLoop).nSeq)
                .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, stMenu(iLoop).nLevel)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, stMenu(iLoop).sParentID)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, " ")                         ' SelectClss
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, stMenu(iLoop).sAddNewClss)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, stMenu(iLoop).sUpdateClss)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, stMenu(iLoop).sDeleteClss)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, stMenu(iLoop).sPrintClss)
                
                .Execute
                
                nSql = nSql + 1
            Next iLoop
        End With
    End If
    
    If nMachineCnt <> 0 Then
        With adoCmd     ' 사용자별 메뉴 입력
            .CommandText = "xp_Person_iPersonMachine"
            .Prepared = True
        
            For iLoop = 0 To nMachineCnt - 1
                Call ClearParameter(adoCmd)
        
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, PersonID)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, stMachine(iLoop).sProcessID)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, stMachine(iLoop).sMachineID)
                
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
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPerson.AddNewPerson"

LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(nCnt + 1)
    With stData
        'S_201312_삼우_99 에 의한 수정-.sOldNNewClss,.sGunMoolMngNo,.sAddress1,.sAddress2,.sAddressAssist 추가
        '                              .sAddress1 -> sAddressJiBun1 로 변경 , .sAddress2 -> sAddressJiBun2 로 변경
        sLog(0) = "xp_Person_iPerson '" & .sPersonID & "', " & .sUserID & ", '" & .sPassword & "', " & .sName & ", '" & _
            .sDepartID & "', '" & .sDutyID & "', '" & .sStartDate & "', '" & _
            .sEndDate & "', '" & .sRegistID & "', '" & .sHandPhone & "', '" & .sPhone & "', '" & .sBirthDay & "', '" & _
            .sSolarClss & "', '" & .sZipCode & "',  '" & .sOldNNewClss & "', '" & .sGunMoolMngNo & "', '" & .sAddress1 & "', '" & _
            .sAddress2 & "', '" & .sAddressAssist & "', '" & .sAddressJiBun1 & "', '" & .sAddressJiBun2 & "', '" & _
            .sRemark & "', '" & .sTeamID & "', '" & .sEName & "','" & .sSMSYN & "'"
    End With
    For iLoop = 1 To nCnt
        With stMenu(iLoop - 1)
            sLog(iLoop) = "xp_Menu_iMenuUser '" & PersonID & "', '" & .sMenuID & "', " & .nSeq & ", " & _
                .nLevel & ", '" & .sParentID & "', '" & .sSelectClss & "', '" & .sAddNewClss & "', '" & .sUpdateClss & "', '" & .sDeleteClss & "', '" & .sUpdateClss & "'"
        End With
    Next iLoop

    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewPerson = False
''        Err.Raise Err.Number, "CPerson.AddNewPerson", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        AddNewPerson = True
    End If


End Function


Public Function UpdatePerson(stData As TPerson, stMenu() As TUSERMENU, nCnt As Integer, _
                                stMachine() As TPersonMachine, nMachineCnt As Integer) As Boolean
    Dim iLoop%, nSql%
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_uPerson"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, stData.sPersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, stData.sUserID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sPassword)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, stData.sName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, stData.sDepartID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, stData.sDutyID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sStartDate)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sEndDate)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 13, stData.sRegistID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 13, stData.sHandPhone)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, stData.sPhone)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sBirthDay)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, stData.sSolarClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 7, stData.sZipCode)
        'S_201312_삼우_99 에 의한 추가-------------------------------------------------------
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, stData.sOldNNewClss)       '도로명,지번주소 구분 0:도로명, 1:지번
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, stData.sGunMoolMngNo)     '건물관리 고유식별번호
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddress1)      '도로명 주소1
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddress2)      '도로명 주소2
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddressAssist)    '도로명 보조 주소
        '----------------------------------------------------------------------------------------
        'S_201312_삼우_99 에 의한 수정(OLD:stData.sAddress1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sAddressJiBun1)
        'S_201312_삼우_99 에 의한 수정(OLD:stData.sAddress2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, stData.sAddressJiBun2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sEMail)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sRemark)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, stData.sTeamID)              '작업조
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, stData.sEName)           '영문이름
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, stData.sSMSYN)                '문자전송 대상여부

        .Execute
    End With
    
    nSql = 1
    
    With adoCmd     '사용자별 메뉴 삭제
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Menu_dUserMenu"
    
        Call ClearParameter(adoCmd)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sPersonID)
        
        .Execute
    
    End With
    
    nSql = nSql + 1
    
    If nCnt <> 0 Then
        With adoCmd     ' 새로운 사용자별 메뉴 입력
            .CommandText = "xp_Menu_iUserMenu"
            .Prepared = True
        
            For iLoop = 0 To nCnt - 1
                Call ClearParameter(adoCmd)
        
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, stMenu(iLoop).sPersonID)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, stMenu(iLoop).sMenuID)
                .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, stMenu(iLoop).nSeq)
                .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, stMenu(iLoop).nLevel)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, stMenu(iLoop).sParentID)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, " ")                         ' SelectClss
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, stMenu(iLoop).sAddNewClss)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, stMenu(iLoop).sUpdateClss)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, stMenu(iLoop).sDeleteClss)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 1, stMenu(iLoop).sPrintClss)
                
            
                .Execute
                nSql = nSql + 1
            Next iLoop
        End With
    End If
    
     With adoCmd     '사용자별 관리공정 내역 삭제
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_dPersonMachine"
    
        Call ClearParameter(adoCmd)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sPersonID)
        
        .Execute
    
    End With
    
    nSql = nSql + 1
    
    If nMachineCnt <> 0 Then
        With adoCmd     ' 사용자별 메뉴 입력
            .CommandText = "xp_Person_iPersonMachine"
            .Prepared = True
        
            For iLoop = 0 To nMachineCnt - 1
                Call ClearParameter(adoCmd)
        
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, stMachine(iLoop).sPersonID)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, stMachine(iLoop).sProcessID)
                .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, stMachine(iLoop).sMachineID)
                
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
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPerson.UpdatePerson"

LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(nCnt + 2)
    With stData
        'S_201312_삼우_99 에 의한 수정-.sOldNNewClss,.sGunMoolMngNo,.sAddress1,.sAddress2,.sAddressAssist 추가
        '                              .sAddress1 -> sAddressJiBun1 로 변경 , .sAddress2 -> sAddressJiBun2 로 변경
        sLog(0) = "xp_Person_uPerson '" & .sPersonID & "', " & .sUserID & ", '" & .sPassword & "', " & .sName & ", '" & _
            .sDepartID & "', '" & .sDutyID & "', '" & .sStartDate & "', '" & _
            .sEndDate & "', '" & .sRegistID & "', '" & .sHandPhone & "', '" & .sPhone & "', '" & .sBirthDay & "', '" & _
            .sSolarClss & "', '" & .sZipCode & "', '" & .sOldNNewClss & "', '" & .sGunMoolMngNo & "', '" & .sAddress1 & "', '" & _
            .sAddress2 & "', '" & .sAddressAssist & "', '" & .sAddressJiBun1 & "', '" & .sAddressJiBun2 & "', '" & _
            .sRemark & "', '" & .sTeamID & "', '" & .sEName & "','" & .sSMSYN & "'"
    End With
    
    sLog(1) = "xp_Person_dMenuUser '" & stData.sPersonID & "'"
    
    For iLoop = 2 To nCnt + 1
        With stMenu(iLoop - 2)
            sLog(iLoop) = "xp_Menu_iUserMenu '" & .sPersonID & "', '" & .sMenuID & "', " & .nSeq & ", " & _
                .nLevel & ", '" & .sParentID & "', '" & .sSelectClss & "', '" & .sAddNewClss & "', '" & .sUpdateClss & "', '" & .sDeleteClss & "', '" & .sUpdateClss & "'"
        End With
    Next iLoop

    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdatePerson = False
''        Err.Raise Err.Number, "CPerson.UpdatePerson", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        UpdatePerson = True
    End If

End Function

Public Function DeletePerson(sPersonID As String, EndDate As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '사용자삭제
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_dPerson"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, sPersonID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, EndDate)
        
        
        .Execute
    
    End With
    
    With adoCmd     '사용자 메뉴삭제
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_dMenuUser"
    
        Call ClearParameter(adoCmd)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, sPersonID)
        
        .Execute
    
    End With

    nSql = 2
    
    g_adoCon.CommitTrans

    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPerson.DeletePerson"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(1)

    sLog(0) = "xp_Code_dPerson '" & sPersonID & "'"
    sLog(1) = "xp_Code_dMenuUser '" & sPersonID & "'"
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeletePerson = False
''        Err.Raise Err.Number, "CPerson.DeletePerson", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        DeletePerson = True
    End If

End Function

Public Function ChangePassWord(ByVal sPersonID As String, ByVal sPassword As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '사용자삭제
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_uPassWord"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, sPersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, sPassword)
        
        .Execute
    
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPerson.ChangePassWord"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    sLog(0) = "xp_Code_uPassWord '" & sPersonID & "', '" & sPassword & "'"
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        ChangePassWord = False
''        Err.Raise Err.Number, "CPerson.UpdatePassWord", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        ChangePassWord = True
    End If

End Function

' 염색스케쥴 프로그램에서 ListBox(염색작업자) 채우기
Public Function GetWorkerList(sDepartID As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_sWorker"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sDepartID)
        
        Set GetWorkerList = .Execute
    End With

    Set adoCmd = Nothing
End Function

' 염색스케쥴 프로그램에서 ListBox(염색작업자) 채우기
Public Function GetWorkTeam() As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_sTeam"

        Set GetWorkTeam = .Execute
    End With

    Set adoCmd = Nothing
End Function


