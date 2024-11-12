VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************
'** System 명 : MRRPLUS2
'** Author    : Wizard
'** 작성자    :
'** 내용      : 거래처 등록
'** 생성일자  :
'** 변경일자  : 2013.12.12
'**------------------------------------------------------------------------------------------------
'
'  요청사항 ID: S_201312_삼우_99
'  요청자:
'  변경날짜 : 2013.12.12
'  작업자   : 오승욱
'  요청내용 : 지번주소에서 도로명 주소로 입력가능하게
'  변경내용 : 창운염직 소스 복사
'**************************************************************************************************
Option Explicit

Type TCustom
    sCustomID       As String '[1] 거래처 코드
    sKCustom        As String '[2] 한글 상호
    sShortCustom    As String '[3] 한글 상호 단축
    sECustom        As String '[4] 영문상호
    sCustomNo       As String '[5] 사업자 번호
    sChief          As String '[6] 대표자 성명
    sCondition      As String '[7] 업태
    sCategory       As String '[8] 업종
    sZipCode        As String '[9] 우편번호
    'S_201312_삼우_99 에 의한 추가---------------------------------------------------
    sOldNNewClss        As String '[10]  도로명,지번주소 구분 0:도로명, 1:지번
    sGunMoolMngNo        As String '[11]  건물관리 고유식별번호
    sAddress1       As String '[12] 도로명주소1
    sAddress2       As String '[13] 도로명주소2
    sAddressAssist       As String '[14] 도로명보조주소
    '--------------------------------------------------------------------------------
    'S_201312_삼우_99 에 의한 수정(OLD:sAddress1)
    sAddressJiBun1       As String '[15] 지번주소1
    'S_201312_삼우_99 에 의한 수정(OLD:sAddress1)
    sAddressJiBun2       As String '[16] 지번주소2
    
    sPhone1         As String '[17] 전화번호1
    sPhone2         As String '[18] 전화번호2
    sFaxNO          As String '[19] 팩스
    sEMail          As String '[20] 전자우편
    sHomePage       As String '[21] 홈 페이지
    sName           As String '[22] 업체 담당자
    sPhone          As String '[23] 담당자 전화번호
    sTradeID        As String '[24] 거래 구분
    sUserID         As String '[25] 웹 로그인
    sUserPassword   As String '[26] 웹 로그 패스워드
    sLossClss       As String '[27] 축율/Loss 포함여부
    sSpendingClss   As String '[28] 소요량 정산방법
    sWorkingClss    As String '[29] 가공료 정산방법
    sCalcClss       As String '[30] Meter->Yard 환산방법
    sPointClss      As String '[31] 소수점 관리방법
End Type

Type TCustomUnit
    sCustomID       As String   ' 거래처 관리 코드
    sArticleID      As String   ' 품명구분 코드
    sStuffWidthID   As String   ' 가공폭
    sWorkID         As String   ' 가공구분 코드
    nUnitPrice      As Single   ' 단가
    nChunkRate      As Single   ' 축율( Loss포함 )
    sETC            As String   ' 비고
End Type


Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetCustom(Optional sCustom As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Custom_sCustom"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, sCustom)
        Set GetCustom = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetCustomOne(sCustomID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Custom_sCustomOne"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        
        Set GetCustomOne = .Execute
    End With
    
    Set adoCmd = Nothing
End Function


Public Function AddNewCustom(TItem As TCustom) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '거래처 입력
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Custom_iCustom"
        
        .Parameters.Append .CreateParameter(, adChar, adParamOutput, 4, "0000")
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sKCustom)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sShortCustom)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sECustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem.sCustomNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TItem.sChief)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, TItem.sCondition)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, TItem.sCategory)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 7, TItem.sZipCode)
        'S_201312_삼우_99 에 의한 추가-------------------------------------------------------
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, TItem.sOldNNewClss)       '도로명,지번주소 구분 0:도로명, 1:지번
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sGunMoolMngNo)     '건물관리 고유식별번호
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, TItem.sAddress1)      '도로명 주소1
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, TItem.sAddress2)      '도로명 주소2
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, TItem.sAddressAssist)    '도로명 보조 주소
        '----------------------------------------------------------------------------------------
        'S_201312_삼우_99 에 의한 수정(OLD:tItem.sAddress1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sAddressJiBun1)
        'S_201312_삼우_99 에 의한 수정(OLD:tItem.sAddress2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, TItem.sAddressJiBun2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, TItem.sPhone1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, TItem.sPhone2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, TItem.sFaxNO)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, TItem.sEMail)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, TItem.sHomePage)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TItem.sName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, TItem.sPhone)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sTradeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sUserID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sUserPassword)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sLossClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sSpendingClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sWorkingClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sCalcClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sPointClss)
        .Execute
        
        TItem.sCustomID = .Parameters(0).Value
    
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTranss
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCustom.AddNewCustom"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With TItem
        'S_201312_삼우_99 에 의한 수정-.sOldNNewClss,.sGunMoolMngNo,.sAddress1,.sAddress2,.sAddressAssist 추가
        '                              .sAddress1 -> sAddressJiBun1 로 변경 , .sAddress2 -> sAddressJiBun2 로 변경
        sLog(0) = "xp_Custom_iCustom '" & .sCustomID & "', '" & .sKCustom & "', '" & .sShortCustom & "', '" & .sECustom & "', '" & _
                .sCustomNo & "', '" & .sChief & "', '" & .sCondition & "', '" & .sCategory & "', '" & _
                .sZipCode & "', , '" & .sOldNNewClss & "', '" & .sGunMoolMngNo & "', '" & .sAddress1 & "', '" & .sAddress2 & "', '" & _
                .sAddressAssist & "', '" & .sAddressJiBun1 & "', '" & .sAddressJiBun2 & "',  '" & .sPhone1 & "', '" & _
                .sPhone2 & "', '" & .sFaxNO & "', '" & .sEMail & "', '" & .sHomePage & "', '" & _
                .sName & "', '" & .sPhone & "', '" & .sTradeID & "', '" & .sUserID & "', '" & _
                .sUserPassword & "', '" & .sLossClss & "', '" & .sSpendingClss & "', '" & .sWorkingClss & "', '" & _
                .sCalcClss & "', '" & .sPointClss & "'"
    End With
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewCustom = False
''        Err.Raise Err.Number, "CCustom.AddNewCustom", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        AddNewCustom = True
    End If


End Function

Public Function UpdateCustom(TItem As TCustom) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '거래처 수정
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Custom_uCustom"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sKCustom)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sShortCustom)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sECustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, TItem.sCustomNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TItem.sChief)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, TItem.sCondition)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, TItem.sCategory)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 7, TItem.sZipCode)
        'S_201312_삼우_99 에 의한 추가-------------------------------------------------------
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, TItem.sOldNNewClss)       '도로명,지번주소 구분 0:도로명, 1:지번
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sGunMoolMngNo)     '건물관리 고유식별번호
        .Parameters.Append .CreateParameter(, adVarWChar, adParamInput, 80, TItem.sAddress1)      '도로명 주소1
        .Parameters.Append .CreateParameter(, adVarWChar, adParamInput, 80, TItem.sAddress2)      '도로명 주소2
        .Parameters.Append .CreateParameter(, adVarWChar, adParamInput, 80, TItem.sAddressAssist)    '도로명 보조 주소
        '----------------------------------------------------------------------------------------
        'S_201312_삼우_99 에 의한 수정(OLD:tItem.sAddress1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sAddressJiBun1)
        'S_201312_삼우_99 에 의한 수정(OLD:tItem.sAddress2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, TItem.sAddressJiBun2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, TItem.sPhone1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, TItem.sPhone2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, TItem.sFaxNO)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, TItem.sEMail)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, TItem.sHomePage)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, TItem.sName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 14, TItem.sPhone)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sTradeID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sUserID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sUserPassword)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sLossClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sSpendingClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sWorkingClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sCalcClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sPointClss)
        
        .Execute
    
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCustom.UpdateCustom"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With TItem
        'S_201312_삼우_99 에 의한 수정-.sOldNNewClss,.sGunMoolMngNo,.sAddress1,.sAddress2,.sAddressAssist 추가
        '                              .sAddress1 -> sAddressJiBun1 로 변경 , .sAddress2 -> sAddressJiBun2 로 변경
        sLog(0) = "xp_Custom_uCustom '" & .sCustomID & "', '" & .sKCustom & "', '" & .sShortCustom & "', '" & .sECustom & "', '" & _
                .sCustomNo & "', '" & .sChief & "', '" & .sCondition & "', '" & .sCategory & "', '" & _
                .sZipCode & "', '" & .sOldNNewClss & "', '" & .sGunMoolMngNo & "', '" & .sAddress1 & "', '" & .sAddress2 & "', '" & _
                .sAddressAssist & "', '" & .sAddressJiBun1 & "', '" & .sAddressJiBun2 & "',  '" & .sPhone1 & "', '" & _
                .sPhone2 & "', '" & .sFaxNO & "', '" & .sEMail & "', '" & .sHomePage & "', '" & _
                .sName & "', '" & .sPhone & "', '" & .sTradeID & "', '" & .sUserID & "', '" & _
                .sUserPassword & "', '" & .sLossClss & "', '" & .sSpendingClss & "', '" & .sWorkingClss & "', '" & _
                .sCalcClss & "', '" & .sPointClss & "'"
    End With
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdateCustom = False
''        Err.Raise Err.Number, "CCustom.UpdateCustom", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        UpdateCustom = True
    End If

End Function

Public Function DeleteCustom(NewKey As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '거래처 삭제
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Custom_dCustom"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewKey)
        
        .Execute
    
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCustom.DeleteCustom"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)
    
    sLog(0) = "xp_Custom_dCustom '" & NewKey & "'"
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeleteCustom = False
''        Err.Raise Err.Number, "CCustom.DeleteCustom", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        DeleteCustom = True
    End If

End Function


Public Function GetCustomUnit(sCustomID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Custom_sCustomUnit"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)

        Set GetCustomUnit = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function DeleteCustomUnit(TItem As TCustomUnit) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '거래처 단가 입력
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Custom_dCustomUnit"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sStuffWidthID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sWorkID)
        
        .Execute
        
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCustom.DeleteCustomUnit"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

''    With TItem
'''        sLog(0) = "xp_Custom_iCustomUnit '" & .sCustomID & "', '" & .sArticleID & "', '" & .sColorID & "', '" & _
''                .sWorkID & "', '" & .sPriceClss & "', '" & .sUnitClss & "', '" & .nPrice
''    End With
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeleteCustomUnit = False
''        Err.Raise Err.Number, "CCustom.DeleteCustomUnit", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        DeleteCustomUnit = True
    End If

End Function

Public Function AddNewCustomUnit(TItem As TCustomUnit) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     '거래처 단가 입력
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Custom_iuCustomUnit"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, TItem.sStuffWidthID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sWorkID)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nUnitPrice)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, TItem.nChunkRate)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sETC)
        
        .Execute
        
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CCustom.AddNewCustomUnit"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With TItem
'        sLog(0) = "xp_Custom_iCustomUnit '" & .sCustomID & "', '" & .sArticleID & "', '" & .sColorID & "', '" & _
                .sWorkID & "', '" & .sPriceClss & "', '" & .sUnitClss & "', '" & .nPrice
    End With
    
    If bError Then  ' 에러 로그
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewCustomUnit = False
''        Err.Raise Err.Number, "CCustom.AddNewCustomUnit", Err.Description
    Else   ' 로그
        Call LogService(sLog)

        AddNewCustomUnit = True
    End If


End Function


Public Function UpdateCustomUnit(TItem As TCustomUnit) As Boolean
'    Dim adoCmd As ADODB.Command
'    Dim bError As Boolean
'    Dim sLog() As String
'    Dim nSql%
'
'    On Error GoTo ErrHandler
'
'    g_adoCon.BeginTrans
'    Set adoCmd = New ADODB.Command
'
'    With adoCmd     '거래처 단가 입력
'        .ActiveConnection = g_adoCon
'        .CommandType = adCmdStoredProc
'        .CommandText = "xp_Custom_uCustomUnit"
'
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sCustomID)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sArticleID)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 5, TItem.sColorID)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, TItem.sWorkID)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sPriceClss)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, TItem.sUnitClss)
'        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, TItem.nPrice)
'
'        .Execute
'
'    End With
'    nSql = 1
'
'    g_adoCon.CommitTrans
'
'    GoTo LogMessage
'
'ErrHandler:
'    g_adoCon.RollbackTrans
'    bError = True
'
'LogMessage:
'    Set adoCmd = Nothing
'
'    ReDim sLog(0)
'
'    With TItem
'        sLog(0) = "xp_Custom_uCustomUnit '" & .sCustomID & "', '" & .sArticleID & "', '" & .sColorID & "', '" & _
'                .sWorkID & "', '" & .sPriceClss & "', '" & .sUnitClss & "', '" & .nPrice
'    End With
'
'    If bError Then  ' 에러 로그
'        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)
'
'        UpdateCustomUnit = False
'        Err.Raise Err.Number, "CCustom.UpdateCustomUnit", Err.Description
'    Else   ' 로그
'        Call LogService(sLog)
'
'        UpdateCustomUnit = True
'    End If


End Function


