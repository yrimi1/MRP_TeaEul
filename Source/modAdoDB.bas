Attribute VB_Name = "modAdoDB"
'**************************************************************************************************
'** System �� : MRRPLUS2
'** Author    : Wizard
'** �ۼ���    : �̰��
'** ����      : DB�� ���� ó���� ���� Module �̴�
'** ��������  : 2012.03.06
'**------------------------------------------------------------------------------------------------
'** ��������    ������  ���泻��
'**************************************************************************************************

Option Explicit
'---------------------------------------
'����Ÿ ���̽� ���� ��� ADO
'---------------------------------------
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameter As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
        
Public giDBConTryCount As Integer           'S_201101_���_02 �� ���� �߰�

'DB Open
Public Function MyOpenDB(Conn As ADODB.Connection, DBPath As String, Optional DBName As String = "", Optional DBUser As String = "", Optional DBPass As String = "", Optional bReadOnly As Boolean = False) As Boolean
On Error GoTo ErrHandler


'Provider=Microsoft.Jet.OLEDB.4.0; Data Source= MDB������ ����������ü���[;
'Jet OLEDB:System Database=�۾��׷����������ǰ�ο������̸�;Jet OLEDB:Registry
'Path=Jet����������Ʈ��Ű;Jet OLEDB:Database Password=��ȣ]
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
    MsgBox "DB ���� ����" & vbCrLf _
            & Err.Number & " : " & Err.Description, vbExclamation, "����"
    MyOpenDB = False

'    If MsgBox("MDAC(Microsoft Data Access Components)�� ��ġ �Ǿ� ���� �ʽ��ϴ�. " & vbCrLf _
'        & "��� Ȩ������ ��Ÿ �ڷ�� 10��[����Ÿ����/�Ӵ�ȸ���� ��ġ����]���� �ٿ�ε��Ͽ� ��ġ�ϼ���." & vbCrLf _
'        & "���� �ٿ�ε� �Ͻðڽ��ϱ�?", vbYesNo + vbDefaultButton1 + vbQuestion, "MDAC����") = vbYes Then
'        'ShellExecute Me.hwnd, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, SW_SHOWNORMAL
'        ShellExecute 0, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, 1
'    End If

End Function

'------------------
'���ڵ�� OPEN
'------------------
'-------------------------------------------------
'[CursorLocation �Ӽ��� CursorType�� ����]
'-------------------------------------------------
'CursorType/CursorLocation Server  Client
'adOpenForwardOnly          O       X
'adOpenKeyset               O       X
'adOpenDynamic              O       X
'adOpenStatic               O       O
'-------------------------------------------------
'[RecordCount �Ǵ� AbsolutePosition�� ���� �Ӽ� ��밡�ɿ���]
'-------------------------------------------------
'CursorType/CursorLocation Server Client
'adOpenForwardOnly          X       X
'adOpenKeyset               O       X
'adOpenDynamic              X       X
'adOpenStatic               O       O
'-------------------------------------------------
'adOpenForwardOnly ����Ʈ, Forward-only Ŀ��. Static Ŀ���� �����ϸ� ���� ���� ���ڵ�θ� �̵��� �� �ִ� Ŀ���̴�. �� Ÿ���� Recordset ��ü�� ������ �� ���� ���ڵ忡 ���� �����͸� �������� �����ǹǷ� �ٸ� ������ Ŀ�� ���� �����Ǵ� �ӵ��� ������.
'adOpenKeyset Keyset Ŀ��. Recordset ��ü�� ������ �Ŀ� �ٸ� ����ڿ� ���ؼ� �߰��ǰų� ������ ���븸 �ݿ����� ���ϸ�, ����� ������ �ݿ��Ѵ�.
'adOpenDynamic Dynamic Ŀ��. Recordset ��ü�� ������ �Ŀ� �ٸ� ����ڿ� ���ؼ� �߰�, ����, ������ ������ �ݿ��ϸ�, Recordset ��ü�� ���� ��� �̵� ������ ����Ѵ�. ��, Provider�� Bookmark�� �������� ���ϴ� ��쿡�� Bookmark�� �������� �ʴ´�.
'adOpenStatic Static Ŀ��. �����ͺ��̽��� �ִ� ���ڵ���� ������ ���纻�� �����ϴ� Ŀ���̴�. �� ���ڵ带 �������� ������ �����͸� ������ �ֱ� ������ ���ڵ带 ������ �Ŀ� �ٸ� ����ڿ� ���ؼ� �߰�, ����, ������ ������ �ݿ����� �ʴ´�.
'-------------------------------------------------
'ó���ӵ�: adOpenForwardOnly > adOpenDynamic > adOpenStatic > adOpenKeyset
'-------------------------------------------------
'locktype: adLockReadOnly,adLockOptimistic

Public Function MyOpenRS(ByVal Conn As ADODB.Connection, rs As ADODB.Recordset, cSrc As String, Optional curType As CursorTypeEnum = adOpenForwardOnly, Optional lockType As ADODB.LockTypeEnum = adLockReadOnly, Optional opt As CommandTypeEnum = adCmdText) As Boolean

    On Error Resume Next

    MyOpenRS = True

    CloseObj rs

    Set rs = New ADODB.Recordset
    rs.Open cSrc, Conn, curType, lockType, opt

    '������ ���� ����� ó��
    Select Case Err.Number
        Case 0
        Case 91
            Call MsgBox("Data Base MDB�� ������� �ʾҽ��ϴ�." & Chr(13) & cSrc, vbInformation, "MDB ����")
        Case 3011
            Call MsgBox("Jet Data Base Engine�� ���� Ȯ���ϼ���." & Chr(13) & cSrc, vbInformation, "Query")
        Case 3261
            Call MsgBox(Err.Description & Chr(13) & cSrc, vbInformation, "MDB ����")
        Case Else
            Call MsgBox(Err.Description & Chr(13) & cSrc, vbInformation, "ERROR")
    End Select

    If Err.Number <> 0 Then
       MyOpenRS = False
    End If

    Err.Clear

End Function

'DB �� table close
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
'    MyWarning "Adodc DB ���� ����" & vbCrLf _
'            & Err.Number & " : " & Err.Description, "����"
'    AdodcConn = False
'End Function

'���̺��� ���ڵ���� ����
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
'        If MsgBox("MDAC(Microsoft Data Access Components)�� ��ġ �Ǿ� ���� �ʽ��ϴ�. " & vbCrLf _
'            & "��� Ȩ������ ��Ÿ �ڷ�� 10��[����Ÿ����/�Ӵ�ȸ���� ��ġ����]���� �ٿ�ε��Ͽ� ��ġ�Ͻʽÿ�." & vbCrLf _
'            & "���� �ٿ�ε� �Ͻðڽ��ϱ�?", vbYesNo + vbDefaultButton1 + vbQuestion, "MDAC����") = vbYes Then
'            'ShellExecute Me.hwnd, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, SW_SHOWNORMAL
'            ShellExecute 0, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, 1
'        End If
'    'End If
'    Err.Clear
'
'End Function


'S_201101_���_02 �� ���� �߰�
'****************************************************************
'*Description:
'*  ADO�� �̿��Ͽ� Database�� �����ϱ�
'****************************************************************
Public Function Gf_DB_ConnectDB() As Boolean
    Dim sConnect$
    
    On Error GoTo ErrHandler
    
    
    If giDBConTryCount <= 3 Then
    
        If g_adoCon Is Nothing Then
            sConnect = "PROVIDER=SQLOLEDB;INTEGRATED SECURITY=SSPI;DATA SOURCE=" & g_sServer & ";DATABASE=" & g_sDatabase & ";UID=sa;PWD=;"
    
            '--------------------------------------------------------------------
            'ZEngine Connection ��
            '--------------------------------------------------------------------
            'Provider=Microsoft.Jet.OLEDB.4.0; Data Source= MDB������ ����������ü���[;
            'Jet OLEDB:System Database=�۾��׷����������ǰ�ο������̸�;Jet OLEDB:Registry
            'Path=Jet����������Ʈ��Ű;Jet OLEDB:Database Password=��ȣ]
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

    Call ErrorBox(Err.Number, Err.Source, Err.Description, "DB Connection ����", True)

    Gf_DB_ConnectDB = False
End Function


'S_201101_���_02 �� ���� �߰�
Public Function Gf_DB_OpenRS(ByVal Conn As ADODB.Connection, rs As ADODB.Recordset, cSrc As String, Optional curType As CursorTypeEnum = adOpenForwardOnly, Optional lockType As ADODB.LockTypeEnum = adLockReadOnly, Optional opt As CommandTypeEnum = adCmdText) As Boolean
 
    On Error GoTo Err_Rtn

Retry_rtn:
    Gs_DB_CloseRs rs

    Set rs = New ADODB.Recordset
    rs.Open cSrc, Conn, curType, lockType, opt

    '������ ���� ����� ó��
    
    If Err.Number <> 0 Then
       Gf_DB_OpenRS = False
    End If

    Gf_DB_OpenRS = True
    
    Err.Clear
    Exit Function
Err_Rtn:
    'DB ���� ���н� �ڵ� Retry (3ȸ ����)
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

'S_201101_���_02 �� ���� �߰�
'������ �����ϰ� ����� ����� ��ȯ
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
            MsgBox "�б��������� DB�� ���ȱ� ������ ���⸦ �� �� �����ϴ�."
        Else
            MsgBox "Error " & Err.Number & ": " & Err.Description   ' in procedure SQLExecute of Module modMDB_Func"
        End If
    End If
 

End Function



'S_201101_���_02 �� ���� �߰�
 Public Sub Gs_DB_CloseRs(prs As ADODB.Recordset)
    On Error Resume Next
    prs.Close
    Set prs = Nothing
 End Sub


