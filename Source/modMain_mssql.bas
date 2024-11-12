Attribute VB_Name = "modMain_mssql"
'---------------------------------------------------------------------------------------
' 모듈명     : modMSSQL_Func
' 최초작성일 : 2012-03-06
' 개발자     :
' 주요기능   : MDB방식의 자주 사용하는 공통함수모음
' 모듈사용법
' 먼저 DBInit()을 호출하여 사용 준비를 한다.
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameter As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'------------------------------------------------------------
'ODBC등록관련 상수 및 API
'------------------------------------------------------------
Private Const KEY_QUERY_VALUE = &H1
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_DWORD = 4

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long ' Note that If you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
'------------------------------------------------------------

'Private g_adoCon As ADODB.Connection

Public g_bDBReadOnly As Boolean    '읽기전용db여부
Public Function DBInit_MSSQL(DBPath As String, Optional DBName As String = "", Optional DBUser As String = "", Optional DBPass As String = "", Optional bReadOnly As Boolean = False) As Boolean
On Error GoTo ErrHandler

    'MDAC 설치여부 확인
    Dim sMdacVer As String
    Dim iRet As Integer
    Dim sConnect As String
    sMdacVer = GetMDACVer()
    If val(Left(sMdacVer, 3)) < 2 Then
        If GetSetting(g_sAppName, "Config", "MDAC_SKIP") <> "1" Then
            iRet = MsgBox("MDAC(Microsoft Data Access Components)가 설치 되어 있지 않거나 버전이 낮습니다.(현재Ver:" & sMdacVer & ")" & vbCrLf _
                & "MicroSoft 홈페이지에서 MDAC를 다운 설치 하십시요.", "MDAC누락")
            
        End If

    End If

    g_bDBReadOnly = bReadOnly

    Set g_adoCon = New ADODB.Connection
    sConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & DBPath & ";DATABASE=" & DBName & ";UID=sa;PWD=;"
    With g_adoCon
        .CursorLocation = adUseClient
        .ConnectionString = sConnect
'        .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=" & DBPath & ";UID=" & DBUser & ";PWD=" & DBPass & ";DATABASE=" & DBName
        
        
        .ConnectionTimeout = 1
        .Open

    End With

    DBInit_MSSQL = True
    On Error GoTo 0
    Exit Function

ErrHandler:
    MsgBox "DB 연결 실패" & vbCrLf _
            & Err.Number & " : " & Err.Description, vbExclamation, "오류"
    DBInit_MSSQL = False

'    If MsgBox("MDAC(Microsoft Data Access Components)가 설치 되어 있지 않습니다. " & vbCrLf _
'        & "당사 홈페이지 자료실 8번[데이타전송/임대회원용 설치파일]에서 다운로드하여 설치하세요." & vbCrLf _
'        & "지금 다운로드 하시겠습니까?", vbYesNo + vbDefaultButton1 + vbQuestion, "MDAC누락") = vbYes Then
'        ShellExecute 0, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, 1
'    End If

End Function

'DB종료
Public Sub DBTerminate()
    CloseAdo g_adoCon
End Sub

'쿼리 결과의 첫번째필드값을 반환
Public Function GetSQLResult(sSQL As String) As Variant
    Dim adoRs As ADODB.Recordset

    Set adoRs = g_adoCon.Execute(sSQL)

    If adoRs.EOF = False Then
        If IsNull(adoRs(0)) Then
            GetSQLResult = ""

        Else
            GetSQLResult = adoRs(0)

        End If

    Else
        GetSQLResult = ""

    End If

    CloseAdo adoRs

End Function

'필드값중 최대길이를 구함
Public Function GetSQLMaxLen(sTBL As String, sFLD As String, Optional sWHERE As String = "") As Integer
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT MAX(LEN(" & sFLD & ")) FROM " & sTBL
    If sWHERE <> "" Then
        sSQL = sSQL & " WHERE " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)

    If IsNull(adoRs(0)) Then
        GetSQLMaxLen = 0

    Else
        GetSQLMaxLen = adoRs(0)

    End If

    CloseAdo adoRs

End Function

'쿼리를 수행하고 적용된 행수를 반환
Public Function SQLExecute(sSQL As String, Optional Silent As Boolean = False) As Long
    Dim adoRs As ADODB.Recordset
    Dim nAffectRow As Long

    On Error GoTo SQLExecute_Error

    Set adoRs = g_adoCon.Execute(sSQL, nAffectRow)
    SQLExecute = nAffectRow

    CloseAdo adoRs

    On Error GoTo 0
    Exit Function

SQLExecute_Error:

    If Silent = False Then
        If g_bDBReadOnly And Err.Number = -2147467259 Then
            MsgBox "읽기전용으로 DB가 열렸기 때문에 쓰기를 할 수 없습니다."
        Else
            MsgBox "Error " & Err.Number & ": " & Err.Description  ' in procedure SQLExecute of Module modMDB_Func"
        End If

    End If

    Resume Next

End Function

'특정테이블의 레코드 수를 반환
Public Function GetSQLCount(sTBL As String, Optional sWHERE As String = "") As Long
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT COUNT(*) FROM " & sTBL
    If sWHERE <> "" Then
        sSQL = sSQL & " WHERE " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)
    GetSQLCount = adoRs(0)

    CloseAdo adoRs

End Function

'순번필드(숫자값)의 최대값을 구함
Public Function GetMAXSEQNum(sTBL As String, sField As String, Optional sWHERE As String = "") As Long
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTBL
    If sWHERE <> "" Then
        sSQL = sSQL & " WHERE " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)

    If IsNull(adoRs(0)) Then
        GetMAXSEQNum = 0
    Else
        GetMAXSEQNum = adoRs(0)
    End If

    CloseAdo adoRs

End Function

'순번필드(문자값)의 최대값을 구함
Public Function GetMAXSEQStr(sTBL As String, sField As String, Optional sWHERE As String = "") As String
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTBL
    If sWHERE <> "" Then
        sSQL = sSQL & " WHERE " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)

    If IsNull(adoRs(0)) Then
        GetMAXSEQStr = ""

    Else
        GetMAXSEQStr = adoRs(0)

    End If

    CloseAdo adoRs

End Function

'순번필드(숫자값)의 다음값을 구함
Public Function GetNextSEQNum(sTBL As String, sField As String, Optional sWHERE As String = "") As Long
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTBL
    If sWHERE <> "" Then
        sSQL = sSQL & " WHERE " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)

    If IsNull(adoRs(0)) Then
        GetNextSEQNum = 1
    Else
        GetNextSEQNum = adoRs(0) + 1
    End If

    CloseAdo adoRs

End Function

'순번필드(문자값)의 다음값을 구함
Public Function GetNextSEQStr(sTBL As String, sField As String, iFieldLen As Integer, Optional sWHERE As String = "") As String
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTBL

    '동일한 길이내에서 구함
    sSQL = sSQL & " WHERE LEN(" & sField & ") = " & iFieldLen

    If sWHERE <> "" Then
        sSQL = sSQL & " AND " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)

    If IsNull(adoRs(0)) Then
        GetNextSEQStr = Format(1, String(iFieldLen, "0"))
    Else
        GetNextSEQStr = Format(adoRs(0) + 1, String(iFieldLen, "0"))
    End If

    CloseAdo adoRs

End Function

'순번필드(문자값)의 부분적으로 다음값을 구함(예: AA001, AA002 ...의 AA에 대한 다음 코드)
Public Function GetNextSEQPart(sTBL As String, sField As String, iFieldLen As Integer, sText As String, Optional sWHERE As String = "") As String
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String
    Dim iTextLen As String
    Dim iSpaceLen As String '필드에서 문자를 제외한 길이
    Dim sResult As String

    GetNextSEQPart = ""

    iTextLen = Len(sText)

    '필드와 데이터의 길이가 같은 경우, 데이터 자체를 반환
    If iFieldLen <= iTextLen Then
        GetNextSEQPart = sText
        Exit Function
    End If

    iSpaceLen = iFieldLen - iTextLen

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTBL

    If iTextLen = 0 Then
        '빈문자열인 경우, 숫자형 데이터중에서 최대값을 구함
        sSQL = sSQL & " WHERE " & sField & " BETWEEN '" & String(iSpaceLen, "0") & "'"
        sSQL = sSQL & " AND '" & String(iSpaceLen, "9") & "'"
        sSQL = sSQL & " AND LEN(" & sField & ") = " & iSpaceLen

        '숫자형만 걸러줌
        sSQL = sSQL & " AND ISNUMERIC(" & sField & ")"

    Else
        '문자가 있는 경우, 문자를 포함하는 데이터중 최대값을 구함
        'sSQL = sSQL & " WHERE LEFT(" & sField & "," & iTextLen & ") = '" & sText & "'" '왼쪽 문자열을 포함하고 있고
        sSQL = sSQL & " WHERE " & sField & " LIKE '" & sText & "%'" '왼쪽 문자열을 포함하고 있고
        sSQL = sSQL & " AND RIGHT(" & sField & "," & iSpaceLen & ")"
        sSQL = sSQL & " BETWEEN '" & String(iSpaceLen, "0") & "' AND '" & String(iSpaceLen, "9") & "'" '나머지 뒤쪽이 숫자인 경우

    End If

    If sWHERE <> "" Then
        sSQL = sSQL & " AND " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)

    If IsNull(adoRs(0)) Then
        '최초 순번
        sResult = sText & Format(1, String(iSpaceLen, "0"))
    Else
        '존재하는 순번+1
        sResult = sText & Format(Right(adoRs(0), iSpaceLen) + 1, String(iSpaceLen, "0"))
    End If

    CloseAdo adoRs

    If Len(sResult) > iFieldLen Then
        MsgBox "해당 순번이 최대값을 넘어 자동으로 값을 구할 수가 없습니다.", vbInformation
    Else
        GetNextSEQPart = sResult
    End If
End Function

'순번필드(문자값)의 부분적으로 다음값을 구함(예: AA001, AA002 ...의 AA에 대한 다음 코드)
'단, 구하는 부분의 길이가 유동적이지 않고 고정적임.
'물론 이경우도 위 GetNextSEQPart()함수로 처리가능하지만, 좀더 빠른 처리를 위해 따로 함수를 만듦
Public Function GetNextSEQFIXPart(sTBL As String, sField As String, iFieldLen As Integer, sText As String, Optional sWHERE As String = "") As String
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String
    Dim iTextLen As String
    Dim iSpaceLen As String '필드에서 문자를 제외한 길이
    Dim sResult As String

    GetNextSEQFIXPart = ""

    iTextLen = Len(sText)
    iSpaceLen = iFieldLen - iTextLen

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTBL

    '문자가 있는 경우, 문자를 포함하는 데이터중 최대값을 구함
    'sSQL = sSQL & " WHERE LEFT(" & sField & "," & iTextLen & ") = '" & sText & "'" '왼쪽 문자열을 포함하는 데이터로 한정지음
    sSQL = sSQL & " WHERE " & sField & " LIKE '" & sText & "%'" '왼쪽 문자열을 포함하는 데이터로 한정지음

    If sWHERE <> "" Then
        sSQL = sSQL & " AND " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)

    If IsNull(adoRs(0)) Then
        '최초 순번
        sResult = sText & Format(1, String(iSpaceLen, "0"))
    Else
        '존재하는 순번+1
        sResult = sText & Format(Right(adoRs(0), iSpaceLen) + 1, String(iSpaceLen, "0"))
    End If

    CloseAdo adoRs

    If Len(sResult) > iFieldLen Then
        MsgBox "해당 순번이 최대값을 넘어 자동으로 값을 구할 수가 없습니다.", vbInformation
    Else
        GetNextSEQFIXPart = sResult
    End If
End Function

'레코드셋 구하기
'반환값: 쿼리결과 레코드 수

'------------------
'레코드셋 OPEN
'------------------
'-------------------------------------------------
'[CursorLocation 속성과 CursorType의 관계]
'-------------------------------------------------
'CursorType/CursorLocation Server  Client
'adOpenForwardOnly          O       X
'adOpenKeyset               O       X
'adOpenDynamic              O       X
'adOpenStatic               O       O
'-------------------------------------------------
'[RecordCount 또는 AbsolutePosition과 같은 속성 사용가능여부]
'-------------------------------------------------
'CursorType/CursorLocation Server Client
'adOpenForwardOnly          X       X
'adOpenKeyset               O       X
'adOpenDynamic              X       X
'adOpenStatic               O       O
'-------------------------------------------------
'adOpenForwardOnly 디폴트, Forward-only 커서. Static 커서와 동일하며 단지 다음 레코드로만 이동할 수 있는 커서이다. 이 타입은 Recordset 개체가 생성될 때 다음 레코드에 대한 포인터만 가지도록 생성되므로 다른 유형의 커서 보다 생성되는 속도가 빠르다.
'adOpenKeyset Keyset 커서. Recordset 개체가 생성된 후에 다른 사용자에 의해서 추가되거나 삭제된 내용만 반영하지 못하며, 변경된 내용은 반영한다.
'adOpenDynamic Dynamic 커서. Recordset 개체가 생성된 후에 다른 사용자에 의해서 추가, 수정, 삭제된 내용을 반영하며, Recordset 개체를 통한 모든 이동 형식을 허용한다. 단, Provider가 Bookmark를 지원하지 못하는 경우에는 Bookmark를 지원하지 않는다.
'adOpenStatic Static 커서. 데이터베이스에 있는 레코드들의 정적인 복사본을 제공하는 커서이다. 즉 레코드를 가져오는 시점의 데이터를 가지고 있기 때문에 레코드를 가져온 후에 다른 사용자에 의해서 추가, 수정, 삭제된 내용이 반영되지 않는다.
'-------------------------------------------------
'처리속도: adOpenForwardOnly > adOpenDynamic > adOpenStatic > adOpenKeyset
'-------------------------------------------------
'locktype: adLockReadOnly,adLockOptimistic

Public Function GetRecordset(adoRs As ADODB.Recordset, sSQL As String, Optional curType As CursorTypeEnum = adOpenForwardOnly, Optional lockType As ADODB.LockTypeEnum = adLockReadOnly, Optional opt As CommandTypeEnum = adCmdText) As Long

    On Error Resume Next

    CloseAdo adoRs

    Set adoRs = New ADODB.Recordset
    'adoRs.CursorLocation = adUseClient
    adoRs.Open sSQL, g_adoCon, curType, lockType, opt

    '오류에 따른 사용자 처리
    Select Case Err.Number
        Case 0
            '결과 레코드수를 구함
            GetRecordset = adoRs.RecordCount

        Case 91
            Call MsgBox("Data Base MDB가 선언되지 않았습니다." & Chr(13) & sSQL, vbInformation, "MDB 오류")

        Case 3011
            Call MsgBox("Jet Data Base Engine의 식을 확인하세요." & Chr(13) & sSQL, vbInformation, "Query")

        Case 3261
            Call MsgBox(Err.Description & Chr(13) & sSQL, vbInformation, "MDB 오류")

        Case Else
            Call MsgBox(Err.Description & Chr(13) & sSQL, vbInformation, "ERROR")

    End Select

    If Err.Number <> 0 Then
       GetRecordset = -1
    End If

    Err.Clear

End Function

'테이블내 동일자료가 있는지 체크(문자형코드)
'TNmae:테이블명
'FNmae:필드명
'sAnd: 덧붙일 조건
'bIfExistShowMSG:만약 동일자료 발견시 처리할지를 묻는 메시지 박스를 보일지 여부
'msgTitle : 메시지 박스에 보일 제목
'msgData: 메시지 박스에 보일 데이터(품목과 같이 검색할 데이타외에 규격을 메시지에 같이 보이게 할때 사용)
Public Function ExistData_str(TName As String, Fname As String, sData As String, Optional sAnd As String = "", Optional bIfExistShowMSG As Boolean = False, Optional msgTitle As String = "", Optional ByVal msgData As String = "") As Boolean
    Dim sWHERE As String

    sWHERE = Fname & "='" & CnvSQLData(sData) & "'"
    If sAnd <> "" Then sWHERE = sWHERE & " AND " & sAnd

    If GetSQLCount(TName, sWHERE) = 0 Then
        ExistData_str = False
    Else

        If bIfExistShowMSG Then
            If msgData = "" Then msgData = sData
            If MsgBox("『" & msgData & "』는 이미 등록된 " & msgTitle & "입니다." & vbCrLf & vbCrLf _
                      & "그래도, 등록하시겠습니까?", vbYesNo + vbDefaultButton2 + vbQuestion, "동일 " & msgTitle & " 발견") = vbNo Then

                ExistData_str = True

            Else
                ExistData_str = False

            End If
        Else
            ExistData_str = True

        End If
    End If
End Function

'테이블내 동일자료가 있는지 체크(문자형코드), 여러 필드 체크(전화번호등에 활용)
'TNmae:테이블명
'FNmae:필드명
'sAnd: 덧붙일 조건
'bIfExistShowMSG:만약 동일자료 발견시 처리할지를 묻는 메시지 박스를 보일지 여부
'msgTitle : 메시지 박스에 보일 제목
'msgData: 메시지 박스에 보일 데이터(품목과 같이 검색할 데이타외에 규격을 메시지에 같이 보이게 할때 사용)
Public Function ExistData_mtstr(TName As String, Fname As Variant, sData As String, Optional sAnd As String = "", Optional bIfExistShowMSG As Boolean = False, Optional msgTitle As String = "", Optional ByVal msgData As String = "") As Boolean
    Dim sWHERE As String
    Dim i As Integer

    sWHERE = "(" & Fname(0) & "='" & sData & "'"

    For i = 1 To UBound(Fname)
        sWHERE = sWHERE & " OR " & Fname(i) & "='" & sData & "'"
    Next

    sWHERE = sWHERE & ")"

    If sAnd <> "" Then sWHERE = sWHERE & " AND " & sAnd

    If GetSQLCount(TName, sWHERE) = 0 Then
        ExistData_mtstr = False
    Else

        If bIfExistShowMSG Then
            If msgData = "" Then msgData = sData
            If MsgBox("『" & msgData & "』는 이미 등록된 " & msgTitle & "입니다." & vbCrLf & vbCrLf _
                      & "그래도, 등록하시겠습니까?", vbYesNo + vbDefaultButton2 + vbQuestion, "동일 " & msgTitle & " 발견") = vbNo Then

                ExistData_mtstr = True

            Else
                ExistData_mtstr = False

            End If
        Else
            ExistData_mtstr = True

        End If

    End If

End Function

'테이블내 동일자료가 있는지 체크(문자형코드), 여러 필드를 합쳐서 체크(주소등에 활용)
'TNmae:테이블명
'FNmae:필드명
'sAnd: 덧붙일 조건
'bIfExistShowMSG:만약 동일자료 발견시 처리할지를 묻는 메시지 박스를 보일지 여부
'msgTitle : 메시지 박스에 보일 제목
'msgData: 메시지 박스에 보일 데이터(품목과 같이 검색할 데이타외에 규격을 메시지에 같이 보이게 할때 사용)
Public Function ExistData_joinstr(TName As String, Fname As Variant, sData As String, Optional sAnd As String = "", Optional bIfExistShowMSG As Boolean = False, Optional msgTitle As String = "", Optional ByVal msgData As String = "") As Boolean
    Dim sWHERE As String
    Dim i As Integer

    'sWhere = FName(0) & "='" & sData & "'"
    sWHERE = Fname(0)
    For i = 1 To UBound(Fname)
        sWHERE = sWHERE & " & ' ' & " & Fname(i)
    Next

    sWHERE = sWHERE & "='" & sData & "'"

    If sAnd <> "" Then sWHERE = sWHERE & " AND " & sAnd

    If GetSQLCount(TName, sWHERE) = 0 Then
        ExistData_joinstr = False
    Else

        If bIfExistShowMSG Then
            If msgData = "" Then msgData = sData
            If MsgBox("『" & msgData & "』는 이미 등록된 " & msgTitle & "입니다." & vbCrLf & vbCrLf _
                      & "그래도, 등록하시겠습니까?", vbYesNo + vbDefaultButton2 + vbQuestion, "동일 " & msgTitle & " 발견") = vbNo Then

                ExistData_joinstr = True

            Else
                ExistData_joinstr = False

            End If
        Else
            ExistData_joinstr = True

        End If
    End If
End Function

'테이블내 동일자료가 있는지 체크(숫자형코드)
'TNmae:테이블명
'FNmae:필드명
Public Function ExistData_num(TName As String, Fname As String, nData As Long, Optional sAnd As String = "") As Boolean
    Dim sWHERE As String

    sWHERE = Fname & "=" & nData
    If sAnd <> "" Then sWHERE = sWHERE & " AND " & sAnd

    If GetSQLCount(TName, sWHERE) = 0 Then
        ExistData_num = False
    Else
        ExistData_num = True
    End If

End Function

'쿼리문장에서 따옴표등에 대한 처리를 하여 에러를 사전에 방지
Public Function CnvSQLData(sData As String) As String
    'Null캐릭터 제거
    If (InStr(sData, Chr(0)) > 0) Then
        sData = Left(sData, InStr(sData, Chr(0)) - 1)
    End If

    CnvSQLData = Replace(sData, "'", "''")

End Function

Public Sub CloseAdo(obj As Object)
    If Not (obj Is Nothing) Then
        If obj.State = adStateOpen Then obj.Close
        Set obj = Nothing
    End If

End Sub

'추가된 레코드의 ID(자동증가 필드값)
Public Function SQLInsertID() As Long
    Dim adoRs As ADODB.Recordset

    Set adoRs = g_adoCon.Execute("SELECT @@Identity", , adCmdText)

    SQLInsertID = adoRs(0).Value

    CloseAdo adoRs

End Function

'ODBC를 생성(시스템DSN생성)
Public Sub CreateMyODBC(DataSourceName As String, sMDBPath As String, Description As String)

    Dim szDriverName As String

    szDriverName = String(255, Chr(32))

    'is access drivers installed?
    '드라이버설치확인및 드라이버경로를 구함
    If Not checkAccessDriver(szDriverName) Then
        MsgBox "Access ODBC 드라이버가 설치되어 있지않습니다. 프로그램 사용전 먼저 설치되어 있어야 합니다.", vbOK + vbCritical
        Exit Sub
    End If

    'is our dsn exist?

    If checkWantedAccessDSN(DataSourceName) = True Then Exit Sub

    If szDriverName = "" Then
        MsgBox "Can't find access ODBC driver.", vbOK + vbCritical
    Else

        If Not CreateAccessDSN(szDriverName, DataSourceName, sMDBPath) Then
            MsgBox "Can't create database ODBC.", vbOK + vbCritical
        End If

    End If

End Sub

'ODBC 드라이버가 설치되어 있는지 체크
Private Function checkAccessDriver(ByRef szDriverName As String) As Boolean

    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean

    bRes = False

    szKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\Microsoft Access Driver (*.mdb)"
    szKeyName = "Driver"
    szKeyValue = String(255, Chr(32))

    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        szDriverName = szKeyValue
        bRes = True
    Else
        bRes = False
    End If

    checkAccessDriver = bRes
End Function

'레지스트리키 존재여부 확인
Private Function isSZKeyExist(szKeyPath As String, szKeyName As String, _
    ByRef szKeyValue As String) As Boolean
    Dim bRes As Boolean
    Dim lRes As Long
    Dim hKey As Long
    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
    szKeyPath, _
    0&, _
    KEY_QUERY_VALUE, _
    hKey)


    If lRes <> ERROR_SUCCESS Then
        isSZKeyExist = False
        Exit Function
    End If

    lRes = RegQueryValueEx(hKey, _
    szKeyName, _
    0&, _
    REG_SZ, _
    ByVal szKeyValue, _
    Len(szKeyValue))
    RegCloseKey (hKey)

    If lRes <> ERROR_SUCCESS Then
        isSZKeyExist = False
        Exit Function
    End If

    isSZKeyExist = True
End Function

'DSN이 이미 생성되어 있는지 체크
Private Function checkWantedAccessDSN(szWantedDSN As String) As Boolean

    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean

    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
    szKeyName = szWantedDSN
    szKeyValue = String(255, Chr(32))

    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        bRes = True
    Else
        bRes = False
    End If

    checkWantedAccessDSN = bRes

End Function

'DSN생성
Private Function CreateAccessDSN(szDriverName As String, szWantedDSN As String, sDBQ As String) As Boolean

    Dim hKey As Long
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim lKeyValue As Long
    Dim lRes As Long
    Dim lSize As Long
    Dim szEmpty As String

    szEmpty = Chr(0)

    lSize = 4
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\" & _
    szWantedDSN, _
    hKey)

    If lRes <> ERROR_SUCCESS Then
        CreateAccessDSN = False
        Exit Function
    End If

    lRes = RegSetValueExString(hKey, "UID", 0&, REG_SZ, _
    szEmpty, Len(szEmpty))

    szKeyValue = sDBQ
    lRes = RegSetValueExString(hKey, "DBQ", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))

    szKeyValue = szDriverName
    lRes = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))

    szKeyValue = "MS Access;"
    lRes = RegSetValueExString(hKey, "FIL", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))

    lKeyValue = 25
    lRes = RegSetValueExLong(hKey, "DriverId", 0&, REG_DWORD, _
    lKeyValue, 4)

    lKeyValue = 0
    lRes = RegSetValueExLong(hKey, "SafeTransactions", 0&, REG_DWORD, _
    lKeyValue, 4)

    lRes = RegCloseKey(hKey)
    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN & "\Engines\Jet"

    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    szKeyPath, _
    hKey)

    If lRes <> ERROR_SUCCESS Then
        CreateAccessDSN = False
        Exit Function
    End If

    lRes = RegSetValueExString(hKey, "ImplicitCommitSync", 0&, REG_SZ, _
    szEmpty, Len(szEmpty))

    szKeyValue = "Yes"
    lRes = RegSetValueExString(hKey, "UserCommitSync", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))

    lKeyValue = 2048
    lRes = RegSetValueExLong(hKey, "MaxBufferSize", 0&, REG_DWORD, _
    lKeyValue, 4)

    lKeyValue = 5
    lRes = RegSetValueExLong(hKey, "PageTimeout", 0&, REG_DWORD, _
    lKeyValue, 4)

    lKeyValue = 3
    lRes = RegSetValueExLong(hKey, "Threads", 0&, REG_DWORD, _
    lKeyValue, 4)

    lRes = RegCloseKey(hKey)
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", _
    hKey)

    If lRes <> ERROR_SUCCESS Then
        CreateAccessDSN = False
        Exit Function
    End If

    szKeyValue = "Microsoft Access Driver (*.mdb)"
    lRes = RegSetValueExString(hKey, szWantedDSN, 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))

    lRes = RegCloseKey(hKey)
    CreateAccessDSN = True

End Function

'MDAC 버전을 구함
Private Function GetMDACVer() As String
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String

'HKEY_LOCAL_MACHINE\Software\Microsoft\DataAccess\FullInstallVer
'HKEY_LOCAL_MACHINE\Software\Microsoft\DataAccess\Version

    szKeyPath = "Software\Microsoft\DataAccess"
    szKeyName = "FullInstallVer" '"Version"
    szKeyValue = String(255, Chr(32))

    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        GetMDACVer = szKeyValue
    End If

End Function




