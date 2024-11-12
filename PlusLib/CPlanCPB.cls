VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlanCPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




Public g_adoCon As ADODB.Connection

Public g_sUserName As String

Type TPlanCPB
    sJobFlag    As String        '[1] �۾�����  Insert / 'U' -- Update
    sPlanDate   As String        '[2] ��ȹ����
    sProcessID  As String        '[3] �����ڵ�
    sOrderID    As String        '[4] ������ȣ
    sPlanClss   As String        '[5] ����, ����
    sEmerClss   As String        '[6] ���, ����
    sPersonID   As String        '[7] �ۼ��� �ڵ�
    sRemark     As String        '[8] ��ȹ����
    nOrderSeq   As Integer       ' OrderSeq --> Color code
    nQty        As Integer       ' ����(yd)
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


'/********************************************************
' * Description : CPB / Rapid ���� ���԰�ȹ
' * ��       �� : pl_mast ��ȹ���� ��������  select
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    �� ¥        �ۼ���    ����                   �������
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     ������    1.0     �ۼ�
' ********************************************************/
Public Function GetPlanCPBList(ByVal sPlanDate As String, ByVal sProcID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
   
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_CPB_sPlanCPB"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPlanDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcID)
        Set GetPlanCPBList = .Execute
        
    End With

    Set adoCmd = Nothing
End Function

'/********************************************************
' * Name        : AddNewPlanCPB
' * Description : CPB / Rapid ���ڵ� �ű�/���� ó��
' * ��       �� : TPlanCPB ����ü�� �� �Է� �� ó��
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    �� ¥        �ۼ���    ����                   �������
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     ������    1.0     �ۼ�
' ********************************************************/
Public Function GetCPBOrder(sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PlanCPB_sOrder"
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 24, sOrder)
        Set GetCPBOrder = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

'/********************************************************
' * Name        : AddNewPlanCPB_Today
' * Description : CPB / Rapid ���ڵ� �ű�/���� ó��
' * ��       �� : TPlanCPB ����ü�� �� �Է� �� ó��
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    �� ¥        �ۼ���    ����                   �������
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     ������    1.0     �ۼ�
' ********************************************************/
Public Function AddNewPlanCPB_Today(ByVal sToDate As String _
                                    , ByVal sPlanDate As String _
                                    , ByVal sProcessID As String _
                                    , ByVal sOrderIDs As String _
                                    , ByVal nOrderSeq As Integer)
                                        
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd     'ǰ�� �Է�
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_PlanCPB_UpdToDay"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sToDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPlanDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, sOrderIDs)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nOrderSeq)

        .Execute
    End With
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    
LogMessage:
    Set adoCmd = Nothing
    MsgBox ("�ش��ڷᰡ ���� ���ڿ� �̹� �ִ��� Ȯ�� �Ͻʽÿ�")
End Function

Public Function AddNewPlanCPB(NewItem As TPlanCPB) As Boolean
    Dim adoCmd As ADODB.Command
    Dim bError As Boolean
    Dim sLog() As String
    Dim nSql%
    Dim dAffected_int As Integer
    
    On Error GoTo ErrHandler
    
    g_adoCon.BeginTrans
    Set adoCmd = New ADODB.Command
    
    With adoCmd     'ǰ�� �Է�
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_CPB_iuCPB"

        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1, nSql)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.sJobFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.sPlanDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.sProcessID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, NewItem.sOrderID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.sPlanClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.sEmerClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.sPersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 500, NewItem.sRemark)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, NewItem.nOrderSeq)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, NewItem.nQty)

        .Execute

         nSql = .Parameters(0).Value
    End With
    
    If nSql = 1 Then
        bError = False
    End If
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPlanCPB.AddNewPlanCPB"
LogMessage:
    Set adoCmd = Nothing
    
    If bError Then  ' ���� �α�
        AddNewPlanCPB = False
    Else   ' �α�
        AddNewPlanCPB = True
    End If

End Function

Public Function GetPlanCPBView(ByVal sProcessID As String _
                            , ByVal nChkDate As Integer _
                            , ByVal sSDate As String _
                            , ByVal sEDate As String _
                            , ByVal nChkCustomID As Integer _
                            , ByVal sCustomID As String _
                            , ByVal nChkArticleID As Integer _
                            , ByVal sArticleID As String _
                            , ByVal nChkOrder As Integer _
                            , ByVal sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_CPB_sPlanCPBView"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 24, sOrder)

        Set GetPlanCPBView = .Execute
    End With
    
    Set adoCmd = Nothing
End Function
'/********************************************************
' * Name        : DeletePlanCPB
' * Description : CPB / Rapid ���ڵ� ����
' * ��       �� : TPlanCPB ����ü�� �� �Է� �� ó��
' *      (CPlanCPB ->  frmPlanCPB )
' *******************************************************************************
' *    �� ¥        �ۼ���    ����                   �������
' *--------------  --------  ------  --------------------------------------------
' * 2003/10/31     ������    1.0     �ۼ�
' ********************************************************/
Public Function DeletePlanCPB(ByVal sPlanDate As String _
                            , ByVal sProcessID As String _
                            , ByVal sOrderID As String _
                            , ByVal nOrderSeq As Integer) As Boolean
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
        .CommandText = "xp_CPB_dCPB"
        
        .Parameters.Append .CreateParameter(, adInteger, adParamOutput, 1, nSql)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPlanDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcessID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nOrderSeq)
        
        .Execute
        nSql = .Parameters(0).Value
    
    End With
    
    If nSql = 1 Then
        bError = False
    Else
        bError = True
    End If
    
    g_adoCon.CommitTrans
    GoTo LogMessage
    
ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CPlanCPB.DeletePlanCPB"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(0)

    sLog(0) = "xp_CPB_dCPB '" & sPlanDate & "', '" & sProcessID & "', '" & sOrderID & "' "
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeletePlanCPB = False
''        Err.Raise Err.Number, "CPlanCPB.DeletePlanCPB", Err.Description
    Else   ' �α�
        Call LogService(sLog)

        DeletePlanCPB = True
    End If

End Function

