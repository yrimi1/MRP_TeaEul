VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type TArticle
    sArticleID      As String   '[1] ǰ�� �ڵ�
    sArticle        As String   '[2] ǰ��
    sThreaID        As String   '���� ���� �ڵ�
    sStuffWidthID   As String   '������ ���� �ڵ�
    DyeingID        As String   '������ ���� �ڵ�
    Weight          As Integer  '�߷�
    UseClss         As String   '��뱸��
    sMixedRate      As String   'ȥ���� S_201901_���_01 �� ���� �߰�
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

Public Function GetArticleID(sID As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Article_sArticle"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sID)
        
        Set GetArticleID = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

' ǰ�� SELECT
Public Function GetArticle(Optional NewKey, Optional UseClss As Integer = 1) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    If IsMissing(NewKey) Then
        NewKey = ""
    Else
        NewKey = Trim(NewKey)
    End If
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Article_sArticle"
        
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 35, NewKey)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 2, UseClss)
        
        Set GetArticle = .Execute
    End With
    
    Set adoCmd = Nothing
End Function


' ǰ�� SELECT
Public Function GetArticleByName(sName As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Article_sArticleByName"
        
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, sName)
        
        Set GetArticleByName = .Execute
    End With
    
    Set adoCmd = Nothing
End Function




Public Function AddNewArticle(NewItem As TArticle) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     'ǰ�� �Է�
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Article_iArticle"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamOutput, 4, "0000")
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, NewItem.sArticleID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 35, NewItem.sArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewItem.sThreaID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewItem.sStuffWidthID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.DyeingID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewItem.Weight)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.UseClss)
        
        'S_201901_���_01 �� ���� �߰� : ȥ���� �߰�
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, NewItem.sMixedRate)
        
        .Execute
        
        NewItem.sArticleID = .Parameters(0).Value
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CArticle.AddNewArticle"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With NewItem
                sLog(0) = "xp_Article_iArticle '" & .sArticleID & "', '" & .sArticle & "', '" & .sThreaID & "', '" & .sStuffWidthID & "', '" & _
                        .DyeingID & "', " & .Weight & ", '" & .UseClss & "'"
    End With
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewArticle = False
''        Err.Raise Err.Number, "CArticle.AddNewArticle", Err.Description
    Else   ' �α�
        Call LogService(sLog)

        AddNewArticle = True
    End If

End Function

Public Function UpdateArticle(NewItem As TArticle) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     'ǰ�� ������Ʈ
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Article_uArticle"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, NewItem.sArticleID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 35, NewItem.sArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewItem.sThreaID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewItem.sStuffWidthID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.DyeingID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, NewItem.Weight)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.UseClss)
        
        'S_201901_���_01 �� ���� �߰� : ȥ���� �߰�
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, NewItem.sMixedRate)
        
        .Execute
        
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CArticle.UpdateArticle"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    With NewItem
        sLog(0) = "xp_Article_uArticle '" & .sArticleID & "', '" & .sArticle & "', '" & .sThreaID & "', '" & .sStuffWidthID & "', '" & _
                        .DyeingID & "', " & .Weight & ", '" & .UseClss & "' "
    End With
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdateArticle = False
''        Err.Raise Err.Number, "CArticle.UpdateArticle", Err.Description
    Else   ' �α�
        Call LogService(sLog)

        UpdateArticle = True
    End If

End Function

Public Function DeleteArticle(NewKey As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     'ǰ�� ����. DelClss�� 1�� ������Ʈ
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Article_dArticle"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, NewKey)
        
        .Execute
    
    End With
    nSql = 1
    
    g_adoCon.CommitTrans
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CArticle.DeleteArticle"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    sLog(0) = "xp_Article_dArticle '" & NewKey & "'"
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeleteArticle = False
''        Err.Raise Err.Number, "CArticle.DeleteArticle", Err.Description
    Else   ' �α�
        Call LogService(sLog)

        DeleteArticle = True
    End If

End Function
