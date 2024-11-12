Attribute VB_Name = "modMain_mssql"
'---------------------------------------------------------------------------------------
' ����     : modMSSQL_Func
' �����ۼ��� : 2012-03-06
' ������     :
' �ֿ���   : MDB����� ���� ����ϴ� �����Լ�����
' ������
' ���� DBInit()�� ȣ���Ͽ� ��� �غ� �Ѵ�.
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameter As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'------------------------------------------------------------
'ODBC��ϰ��� ��� �� API
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

Public g_bDBReadOnly As Boolean    '�б�����db����
Public Function DBInit_MSSQL(DBPath As String, Optional DBName As String = "", Optional DBUser As String = "", Optional DBPass As String = "", Optional bReadOnly As Boolean = False) As Boolean
On Error GoTo ErrHandler

    'MDAC ��ġ���� Ȯ��
    Dim sMdacVer As String
    Dim iRet As Integer
    Dim sConnect As String
    sMdacVer = GetMDACVer()
    If val(Left(sMdacVer, 3)) < 2 Then
        If GetSetting(g_sAppName, "Config", "MDAC_SKIP") <> "1" Then
            iRet = MsgBox("MDAC(Microsoft Data Access Components)�� ��ġ �Ǿ� ���� �ʰų� ������ �����ϴ�.(����Ver:" & sMdacVer & ")" & vbCrLf _
                & "MicroSoft Ȩ���������� MDAC�� �ٿ� ��ġ �Ͻʽÿ�.", "MDAC����")
            
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
    MsgBox "DB ���� ����" & vbCrLf _
            & Err.Number & " : " & Err.Description, vbExclamation, "����"
    DBInit_MSSQL = False

'    If MsgBox("MDAC(Microsoft Data Access Components)�� ��ġ �Ǿ� ���� �ʽ��ϴ�. " & vbCrLf _
'        & "��� Ȩ������ �ڷ�� 8��[����Ÿ����/�Ӵ�ȸ���� ��ġ����]���� �ٿ�ε��Ͽ� ��ġ�ϼ���." & vbCrLf _
'        & "���� �ٿ�ε� �Ͻðڽ��ϱ�?", vbYesNo + vbDefaultButton1 + vbQuestion, "MDAC����") = vbYes Then
'        ShellExecute 0, "Open", "http://www.mijinsoft.co.kr/download/mdac.exe", &O0, &O0, 1
'    End If

End Function

'DB����
Public Sub DBTerminate()
    CloseAdo g_adoCon
End Sub

'���� ����� ù��°�ʵ尪�� ��ȯ
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

'�ʵ尪�� �ִ���̸� ����
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

'������ �����ϰ� ����� ����� ��ȯ
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
            MsgBox "�б��������� DB�� ���ȱ� ������ ���⸦ �� �� �����ϴ�."
        Else
            MsgBox "Error " & Err.Number & ": " & Err.Description  ' in procedure SQLExecute of Module modMDB_Func"
        End If

    End If

    Resume Next

End Function

'Ư�����̺��� ���ڵ� ���� ��ȯ
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

'�����ʵ�(���ڰ�)�� �ִ밪�� ����
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

'�����ʵ�(���ڰ�)�� �ִ밪�� ����
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

'�����ʵ�(���ڰ�)�� �������� ����
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

'�����ʵ�(���ڰ�)�� �������� ����
Public Function GetNextSEQStr(sTBL As String, sField As String, iFieldLen As Integer, Optional sWHERE As String = "") As String
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTBL

    '������ ���̳����� ����
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

'�����ʵ�(���ڰ�)�� �κ������� �������� ����(��: AA001, AA002 ...�� AA�� ���� ���� �ڵ�)
Public Function GetNextSEQPart(sTBL As String, sField As String, iFieldLen As Integer, sText As String, Optional sWHERE As String = "") As String
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String
    Dim iTextLen As String
    Dim iSpaceLen As String '�ʵ忡�� ���ڸ� ������ ����
    Dim sResult As String

    GetNextSEQPart = ""

    iTextLen = Len(sText)

    '�ʵ�� �������� ���̰� ���� ���, ������ ��ü�� ��ȯ
    If iFieldLen <= iTextLen Then
        GetNextSEQPart = sText
        Exit Function
    End If

    iSpaceLen = iFieldLen - iTextLen

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTBL

    If iTextLen = 0 Then
        '���ڿ��� ���, ������ �������߿��� �ִ밪�� ����
        sSQL = sSQL & " WHERE " & sField & " BETWEEN '" & String(iSpaceLen, "0") & "'"
        sSQL = sSQL & " AND '" & String(iSpaceLen, "9") & "'"
        sSQL = sSQL & " AND LEN(" & sField & ") = " & iSpaceLen

        '�������� �ɷ���
        sSQL = sSQL & " AND ISNUMERIC(" & sField & ")"

    Else
        '���ڰ� �ִ� ���, ���ڸ� �����ϴ� �������� �ִ밪�� ����
        'sSQL = sSQL & " WHERE LEFT(" & sField & "," & iTextLen & ") = '" & sText & "'" '���� ���ڿ��� �����ϰ� �ְ�
        sSQL = sSQL & " WHERE " & sField & " LIKE '" & sText & "%'" '���� ���ڿ��� �����ϰ� �ְ�
        sSQL = sSQL & " AND RIGHT(" & sField & "," & iSpaceLen & ")"
        sSQL = sSQL & " BETWEEN '" & String(iSpaceLen, "0") & "' AND '" & String(iSpaceLen, "9") & "'" '������ ������ ������ ���

    End If

    If sWHERE <> "" Then
        sSQL = sSQL & " AND " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)

    If IsNull(adoRs(0)) Then
        '���� ����
        sResult = sText & Format(1, String(iSpaceLen, "0"))
    Else
        '�����ϴ� ����+1
        sResult = sText & Format(Right(adoRs(0), iSpaceLen) + 1, String(iSpaceLen, "0"))
    End If

    CloseAdo adoRs

    If Len(sResult) > iFieldLen Then
        MsgBox "�ش� ������ �ִ밪�� �Ѿ� �ڵ����� ���� ���� ���� �����ϴ�.", vbInformation
    Else
        GetNextSEQPart = sResult
    End If
End Function

'�����ʵ�(���ڰ�)�� �κ������� �������� ����(��: AA001, AA002 ...�� AA�� ���� ���� �ڵ�)
'��, ���ϴ� �κ��� ���̰� ���������� �ʰ� ��������.
'���� �̰�쵵 �� GetNextSEQPart()�Լ��� ó������������, ���� ���� ó���� ���� ���� �Լ��� ����
Public Function GetNextSEQFIXPart(sTBL As String, sField As String, iFieldLen As Integer, sText As String, Optional sWHERE As String = "") As String
    Dim adoRs As ADODB.Recordset
    Dim sSQL As String
    Dim iTextLen As String
    Dim iSpaceLen As String '�ʵ忡�� ���ڸ� ������ ����
    Dim sResult As String

    GetNextSEQFIXPart = ""

    iTextLen = Len(sText)
    iSpaceLen = iFieldLen - iTextLen

    sSQL = "SELECT MAX(" & sField & ") FROM " & sTBL

    '���ڰ� �ִ� ���, ���ڸ� �����ϴ� �������� �ִ밪�� ����
    'sSQL = sSQL & " WHERE LEFT(" & sField & "," & iTextLen & ") = '" & sText & "'" '���� ���ڿ��� �����ϴ� �����ͷ� ��������
    sSQL = sSQL & " WHERE " & sField & " LIKE '" & sText & "%'" '���� ���ڿ��� �����ϴ� �����ͷ� ��������

    If sWHERE <> "" Then
        sSQL = sSQL & " AND " & sWHERE
    End If

    Set adoRs = g_adoCon.Execute(sSQL)

    If IsNull(adoRs(0)) Then
        '���� ����
        sResult = sText & Format(1, String(iSpaceLen, "0"))
    Else
        '�����ϴ� ����+1
        sResult = sText & Format(Right(adoRs(0), iSpaceLen) + 1, String(iSpaceLen, "0"))
    End If

    CloseAdo adoRs

    If Len(sResult) > iFieldLen Then
        MsgBox "�ش� ������ �ִ밪�� �Ѿ� �ڵ����� ���� ���� ���� �����ϴ�.", vbInformation
    Else
        GetNextSEQFIXPart = sResult
    End If
End Function

'���ڵ�� ���ϱ�
'��ȯ��: ������� ���ڵ� ��

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

Public Function GetRecordset(adoRs As ADODB.Recordset, sSQL As String, Optional curType As CursorTypeEnum = adOpenForwardOnly, Optional lockType As ADODB.LockTypeEnum = adLockReadOnly, Optional opt As CommandTypeEnum = adCmdText) As Long

    On Error Resume Next

    CloseAdo adoRs

    Set adoRs = New ADODB.Recordset
    'adoRs.CursorLocation = adUseClient
    adoRs.Open sSQL, g_adoCon, curType, lockType, opt

    '������ ���� ����� ó��
    Select Case Err.Number
        Case 0
            '��� ���ڵ���� ����
            GetRecordset = adoRs.RecordCount

        Case 91
            Call MsgBox("Data Base MDB�� ������� �ʾҽ��ϴ�." & Chr(13) & sSQL, vbInformation, "MDB ����")

        Case 3011
            Call MsgBox("Jet Data Base Engine�� ���� Ȯ���ϼ���." & Chr(13) & sSQL, vbInformation, "Query")

        Case 3261
            Call MsgBox(Err.Description & Chr(13) & sSQL, vbInformation, "MDB ����")

        Case Else
            Call MsgBox(Err.Description & Chr(13) & sSQL, vbInformation, "ERROR")

    End Select

    If Err.Number <> 0 Then
       GetRecordset = -1
    End If

    Err.Clear

End Function

'���̺� �����ڷᰡ �ִ��� üũ(�������ڵ�)
'TNmae:���̺��
'FNmae:�ʵ��
'sAnd: ������ ����
'bIfExistShowMSG:���� �����ڷ� �߽߰� ó�������� ���� �޽��� �ڽ��� ������ ����
'msgTitle : �޽��� �ڽ��� ���� ����
'msgData: �޽��� �ڽ��� ���� ������(ǰ��� ���� �˻��� ����Ÿ�ܿ� �԰��� �޽����� ���� ���̰� �Ҷ� ���)
Public Function ExistData_str(TName As String, Fname As String, sData As String, Optional sAnd As String = "", Optional bIfExistShowMSG As Boolean = False, Optional msgTitle As String = "", Optional ByVal msgData As String = "") As Boolean
    Dim sWHERE As String

    sWHERE = Fname & "='" & CnvSQLData(sData) & "'"
    If sAnd <> "" Then sWHERE = sWHERE & " AND " & sAnd

    If GetSQLCount(TName, sWHERE) = 0 Then
        ExistData_str = False
    Else

        If bIfExistShowMSG Then
            If msgData = "" Then msgData = sData
            If MsgBox("��" & msgData & "���� �̹� ��ϵ� " & msgTitle & "�Դϴ�." & vbCrLf & vbCrLf _
                      & "�׷���, ����Ͻðڽ��ϱ�?", vbYesNo + vbDefaultButton2 + vbQuestion, "���� " & msgTitle & " �߰�") = vbNo Then

                ExistData_str = True

            Else
                ExistData_str = False

            End If
        Else
            ExistData_str = True

        End If
    End If
End Function

'���̺� �����ڷᰡ �ִ��� üũ(�������ڵ�), ���� �ʵ� üũ(��ȭ��ȣ� Ȱ��)
'TNmae:���̺��
'FNmae:�ʵ��
'sAnd: ������ ����
'bIfExistShowMSG:���� �����ڷ� �߽߰� ó�������� ���� �޽��� �ڽ��� ������ ����
'msgTitle : �޽��� �ڽ��� ���� ����
'msgData: �޽��� �ڽ��� ���� ������(ǰ��� ���� �˻��� ����Ÿ�ܿ� �԰��� �޽����� ���� ���̰� �Ҷ� ���)
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
            If MsgBox("��" & msgData & "���� �̹� ��ϵ� " & msgTitle & "�Դϴ�." & vbCrLf & vbCrLf _
                      & "�׷���, ����Ͻðڽ��ϱ�?", vbYesNo + vbDefaultButton2 + vbQuestion, "���� " & msgTitle & " �߰�") = vbNo Then

                ExistData_mtstr = True

            Else
                ExistData_mtstr = False

            End If
        Else
            ExistData_mtstr = True

        End If

    End If

End Function

'���̺� �����ڷᰡ �ִ��� üũ(�������ڵ�), ���� �ʵ带 ���ļ� üũ(�ּҵ Ȱ��)
'TNmae:���̺��
'FNmae:�ʵ��
'sAnd: ������ ����
'bIfExistShowMSG:���� �����ڷ� �߽߰� ó�������� ���� �޽��� �ڽ��� ������ ����
'msgTitle : �޽��� �ڽ��� ���� ����
'msgData: �޽��� �ڽ��� ���� ������(ǰ��� ���� �˻��� ����Ÿ�ܿ� �԰��� �޽����� ���� ���̰� �Ҷ� ���)
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
            If MsgBox("��" & msgData & "���� �̹� ��ϵ� " & msgTitle & "�Դϴ�." & vbCrLf & vbCrLf _
                      & "�׷���, ����Ͻðڽ��ϱ�?", vbYesNo + vbDefaultButton2 + vbQuestion, "���� " & msgTitle & " �߰�") = vbNo Then

                ExistData_joinstr = True

            Else
                ExistData_joinstr = False

            End If
        Else
            ExistData_joinstr = True

        End If
    End If
End Function

'���̺� �����ڷᰡ �ִ��� üũ(�������ڵ�)
'TNmae:���̺��
'FNmae:�ʵ��
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

'�������忡�� ����ǥ� ���� ó���� �Ͽ� ������ ������ ����
Public Function CnvSQLData(sData As String) As String
    'Nullĳ���� ����
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

'�߰��� ���ڵ��� ID(�ڵ����� �ʵ尪)
Public Function SQLInsertID() As Long
    Dim adoRs As ADODB.Recordset

    Set adoRs = g_adoCon.Execute("SELECT @@Identity", , adCmdText)

    SQLInsertID = adoRs(0).Value

    CloseAdo adoRs

End Function

'ODBC�� ����(�ý���DSN����)
Public Sub CreateMyODBC(DataSourceName As String, sMDBPath As String, Description As String)

    Dim szDriverName As String

    szDriverName = String(255, Chr(32))

    'is access drivers installed?
    '����̹���ġȮ�ι� ����̹���θ� ����
    If Not checkAccessDriver(szDriverName) Then
        MsgBox "Access ODBC ����̹��� ��ġ�Ǿ� �����ʽ��ϴ�. ���α׷� ����� ���� ��ġ�Ǿ� �־�� �մϴ�.", vbOK + vbCritical
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

'ODBC ����̹��� ��ġ�Ǿ� �ִ��� üũ
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

'������Ʈ��Ű ���翩�� Ȯ��
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

'DSN�� �̹� �����Ǿ� �ִ��� üũ
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

'DSN����
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

'MDAC ������ ����
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




