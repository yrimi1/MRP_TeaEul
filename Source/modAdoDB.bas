Attribute VB_Name = "modAdoDB"
'**************************************************************************************************
'** System 명 : MRRPLUS2
'** Author    : Wizard
'** 작성자    : 이경미
'** 내용      : DB의 연결 처리를 위한 Module 이다
'** 생성일자  : 2012.03.06
'**------------------------------------------------------------------------------------------------
'** 변경일자    변경자  변경내용
'**************************************************************************************************

Option Explicit
'---------------------------------------
'데이타 베이스 관련 모듈 ADO
'---------------------------------------
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameter As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
        
Public giDBConTryCount As Integer           'S_201101_대안_02 에 의한 추가

'DB Open
Public Function MyOpenDB(Conn As ADODB.Connection, DBPath As String, Optional DBName As String = "", Optional DBUser As String = "", Optional DBPass As String = "", Optional bReadOnly As Boolean = False) As Boolean
On Error GoTo ErrHandler


'Provider=Microsoft.Jet.OLEDB.4.0; Data Source= MDB파일의 물리적인전체경로[;
'Jet OLEDB:System Database=작업그룹정보파일의경로와파일이름;Jet OLEDB:Registry
'Path=Jet엔진레지스트리키;Jet OLEDB:Database Password=암호]
'ex) Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\MyDb.mdb

'    With Conn
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MdbName
'        If strPWD <> "" Then .ConnectionString = .ConnectionString & ";Jet OLEDB:Database Password=" & strPWD
'        .ConnectionTimeout = 30
'        .CursorLocation = adUseClient
'        .Open
'    End With

    With Conn
        .CursorLocation = adUseClient
        .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=" & DBPath & ";UID=" & DBUser & ";PWD=" & DBPass & ";DATABASE=" & DBName
        .ConnectionTimeout = 1
        .Open

    End With

    MyOpenDB = True
    Exit Function

ErrHandler:
    MsgBox "DB 연결 실패" & vbCrLf _
            & Err.Number & " : " & Err.Description, vbExclamation, "오류"
    MyOpenDB = False

'    If MsgBox("MDAC(Microsoft Data Access Components)가 설치 되어 있지 않습니다. " & vbCrLf _
'        & "당사 홈페이지 기타 자료실 10번[데이타전송/임대회원용 설치파일]에서 다운로드하여 설치하세요." & vbCrLf _
'        & "지금 다운로드 하시겠습니까?", vbYesNo + vbDefaultButton1 + vbQuestion, "MDAC누락") = vbYes Then
'        'ShellExecute Me.hwnd, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, SW_SHOWNORMAL
'        ShellExecute 0, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, 1
'    End If

End Function

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

Public Function MyOpenRS(ByVal Conn As ADODB.Connection, rs As ADODB.Recordset, cSrc As String, Optional curType As CursorTypeEnum = adOpenForwardOnly, Optional lockType As ADODB.LockTypeEnum = adLockReadOnly, Optional opt As CommandTypeEnum = adCmdText) As Boolean

    On Error Resume Next

    MyOpenRS = True

    CloseObj rs

    Set rs = New ADODB.Recordset
    rs.Open cSrc, Conn, curType, lockType, opt

    '오류에 따른 사용자 처리
    Select Case Err.Number
        Case 0
        Case 91
            Call MsgBox("Data Base MDB가 선언되지 않았습니다." & Chr(13) & cSrc, vbInformation, "MDB 오류")
        Case 3011
            Call MsgBox("Jet Data Base Engine의 식을 확인하세요." & Chr(13) & cSrc, vbInformation, "Query")
        Case 3261
            Call MsgBox(Err.Description & Chr(13) & cSrc, vbInformation, "MDB 오류")
        Case Else
            Call MsgBox(Err.Description & Chr(13) & cSrc, vbInformation, "ERROR")
    End Select

    If Err.Number <> 0 Then
       MyOpenRS = False
    End If

    Err.Clear

End Function

'DB 및 table close
Public Sub CloseObj(obj As Object)
    If Not (obj Is Nothing) Then
        If obj.State = adStateOpen Then obj.Close
        Set obj = Nothing
    End If
End Sub

'Public Function AdodcConn(adodc As adodc, MdbName As String, Optional strPWD As String = "") As Boolean
'On Error GoTo ErrHandler
'    With adodc
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MdbName
'        If strPWD <> "" Then .ConnectionString = .ConnectionString & ";Jet OLEDB:Database Password=" & strPWD
'        .ConnectionTimeout = 30
'    End With
'
'    AdodcConn = True
'    Exit Function
'
'ErrHandler:
'    MyWarning "Adodc DB 연결 실패" & vbCrLf _
'            & Err.Number & " : " & Err.Description, "오류"
'    AdodcConn = False
'End Function

'테이블내의 레코드수를 구함
Public Function RecCount(objConn As ADODB.Connection, sTblName As String, Optional sWHERE As String = "") As Long
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String

   On Error GoTo RecCount_Error

    RecCount = 0
    sSQL = "SELECT COUNT(*) FROM " & sTblName & sWHERE
    If MyOpenRS(objConn, adoRs, sSQL) = False Then
        Exit Function
    End If

    RecCount = adoRs(0)

    CloseObj adoRs

   On Error GoTo 0
   Exit Function

RecCount_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RecCount of Module MduAdoDB"
    Resume Next
End Function

'Public Function MyOpenADODB(cObject As ADODB.Connection, sProvider As String) As Boolean
'    Dim cnnString As String
'
'    On Error GoTo ERR_RTN
'
'    Set cObject = New ADODB.Connection
'
'    cnnString = sProvider
'    cObject.ConnectionTimeout = 30
'    cObject.Open cnnString
'
'    If Err.Number = 0 Then
'        MyOpenADODB = True
'    Else
'        MyOpenADODB = False
'        Err.Clear
'    End If
'
'    Exit Function
'
'ERR_RTN:
'    MsgBox Err.Number & " " & Err.Description, vbExclamation
'
'    'If Err.Number = 429 Then
'
'        If MsgBox("MDAC(Microsoft Data Access Components)가 설치 되어 있지 않습니다. " & vbCrLf _
'            & "당사 홈페이지 기타 자료실 10번[데이타전송/임대회원용 설치파일]에서 다운로드하여 설치하십시요." & vbCrLf _
'            & "지금 다운로드 하시겠습니까?", vbYesNo + vbDefaultButton1 + vbQuestion, "MDAC누락") = vbYes Then
'            'ShellExecute Me.hwnd, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, SW_SHOWNORMAL
'            ShellExecute 0, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, 1
'        End If
'    'End If
'    Err.Clear
'
'End Function


'S_201101_대안_02 에 의한 추가
'****************************************************************
'*Description:
'*  ADO를 이용하여 Database에 접속하기
'****************************************************************
Public Function Gf_DB_ConnectDB() As Boolean
    Dim sConnect$
    
    On Error GoTo ErrHandler
    
    
    If giDBConTryCount <= 3 Then
    
        If g_adoCon Is Nothing Then
            sConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sServer & ";DATABASE=" & g_sDatabase & ";UID=sa;PWD=;"
    
            '--------------------------------------------------------------------
            'ZEngine Connection 시
            '--------------------------------------------------------------------
            'Provider=Microsoft.Jet.OLEDB.4.0; Data Source= MDB파일의 물리적인전체경로[;
            'Jet OLEDB:System Database=작업그룹정보파일의경로와파일이름;Jet OLEDB:Registry
            'Path=Jet엔진레지스트리키;Jet OLEDB:Database Password=암호]
            'ex) Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\MyDb.mdb
            
            'With Conn
            '    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MdbName
            '    If strPWD <> "" Then .ConnectionString = .ConnectionString & ";Jet OLEDB:Database Password=" & strPWD
            '    .ConnectionTimeout = 30
            '    .CursorLocation = adUseClient
            '    .Open
            'End With
            '--------------------------------------------------------------------
            
            Set g_adoCon = New ADODB.Connection
            With g_adoCon
                .CommandTimeout = 60
                .ConnectionString = sConnect
                .CursorLocation = adUseClient
                .Open sConnect
            End With
            Gf_DB_ConnectDB = True
        ElseIf g_adoCon.State = adStateOpen Then
            Gf_DB_ConnectDB = True
        Else
            Gf_DB_ConnectDB = False
        End If
        
        giDBConTryCount = giDBConTryCount + 1
        
    End If
    
    Exit Function
ErrHandler:
''''''    Unload frm_cm_Splash

    Call ErrorBox(Err.Number, Err.Source, Err.Description, "DB Connection 실패", True)

    Gf_DB_ConnectDB = False
End Function


'S_201101_대안_02 에 의한 추가
Public Function Gf_DB_OpenRS(ByVal Conn As ADODB.Connection, rs As ADODB.Recordset, cSrc As String, Optional curType As CursorTypeEnum = adOpenForwardOnly, Optional lockType As ADODB.LockTypeEnum = adLockReadOnly, Optional opt As CommandTypeEnum = adCmdText) As Boolean
 
    On Error GoTo Err_Rtn

Retry_rtn:
    Gs_DB_CloseRs rs

    Set rs = New ADODB.Recordset
    rs.Open cSrc, Conn, curType, lockType, opt

    '오류에 따른 사용자 처리
    
    If Err.Number <> 0 Then
       Gf_DB_OpenRS = False
    End If

    Gf_DB_OpenRS = True
    
    Err.Clear
    Exit Function
Err_Rtn:
    'DB 연결 실패시 자동 Retry (3회 까지)
   If Err.Number = -2147467259 And giDBConTryCount <= 3 Then
        Set g_adoCon = Nothing
        If Gf_DB_ConnectDB() = False Then Exit Function
        Set Conn = g_adoCon
        GoTo Retry_rtn
        
    End If
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & "," & Err.Description, vbCritical, "[Gf_DB_OpenRS]"
    End If
    
    
    
    
End Function

'S_201101_대안_02 에 의한 추가
'쿼리를 수행하고 적용된 행수를 반환
Public Function Gs_DB_SqlExecute(ByVal Conn As ADODB.Connection, sSQL As String, Optional Silent As Boolean = False) As Long
    Dim adoRs                           As ADODB.Recordset
    Dim nAffectRow                      As Long
    Dim sLog()                          As String
    
    On Error GoTo SQLExecute_Error

    Set adoRs = Conn.Execute(sSQL, nAffectRow)
    Gs_DB_SqlExecute = nAffectRow
    
    CloseAdo adoRs

    Exit Function

SQLExecute_Error:
    '-2147467259
    If Silent = False Then
        If g_bDBReadOnly And Err.Number = -2147467259 Then
            MsgBox "읽기전용으로 DB가 열렸기 때문에 쓰기를 할 수 없습니다."
        Else
            MsgBox "Error " & Err.Number & ": " & Err.Description   ' in procedure SQLExecute of Module modMDB_Func"
        End If
    End If
 

End Function



'S_201101_대안_02 에 의한 추가
 Public Sub Gs_DB_CloseRs(prs As ADODB.Recordset)
    On Error Resume Next
    prs.Close
    Set prs = Nothing
 End Sub


