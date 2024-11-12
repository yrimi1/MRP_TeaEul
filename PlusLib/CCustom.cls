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
'** System �� : MRRPLUS2
'** Author    : Wizard
'** �ۼ���    :
'** ����      : �ŷ�ó ���
'** ��������  :
'** ��������  : 2013.12.12
'**------------------------------------------------------------------------------------------------
'
'  ��û���� ID: S_201312_���_99
'  ��û��:
'  ���泯¥ : 2013.12.12
'  �۾���   : ���¿�
'  ��û���� : �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�
'  ���泻�� : â��� �ҽ� ����
'**************************************************************************************************
Option Explicit

Type TCustom
    sCustomID       As String '[1] �ŷ�ó �ڵ�
    sKCustom        As String '[2] �ѱ� ��ȣ
    sShortCustom    As String '[3] �ѱ� ��ȣ ����
    sECustom        As String '[4] ������ȣ
    sCustomNo       As String '[5] ����� ��ȣ
    sChief          As String '[6] ��ǥ�� ����
    sCondition      As String '[7] ����
    sCategory       As String '[8] ����
    sZipCode        As String '[9] ������ȣ
    'S_201312_���_99 �� ���� �߰�---------------------------------------------------
    sOldNNewClss        As String '[10]  ���θ�,�����ּ� ���� 0:���θ�, 1:����
    sGunMoolMngNo        As String '[11]  �ǹ����� �����ĺ���ȣ
    sAddress1       As String '[12] ���θ��ּ�1
    sAddress2       As String '[13] ���θ��ּ�2
    sAddressAssist       As String '[14] ���θ������ּ�
    '--------------------------------------------------------------------------------
    'S_201312_���_99 �� ���� ����(OLD:sAddress1)
    sAddressJiBun1       As String '[15] �����ּ�1
    'S_201312_���_99 �� ���� ����(OLD:sAddress1)
    sAddressJiBun2       As String '[16] �����ּ�2
    
    sPhone1         As String '[17] ��ȭ��ȣ1
    sPhone2         As String '[18] ��ȭ��ȣ2
    sFaxNO          As String '[19] �ѽ�
    sEMail          As String '[20] ���ڿ���
    sHomePage       As String '[21] Ȩ ������
    sName           As String '[22] ��ü �����
    sPhone          As String '[23] ����� ��ȭ��ȣ
    sTradeID        As String '[24] �ŷ� ����
    sUserID         As String '[25] �� �α���
    sUserPassword   As String '[26] �� �α� �н�����
    sLossClss       As String '[27] ����/Loss ���Կ���
    sSpendingClss   As String '[28] �ҿ䷮ ������
    sWorkingClss    As String '[29] ������ ������
    sCalcClss       As String '[30] Meter->Yard ȯ����
    sPointClss      As String '[31] �Ҽ��� �������
End Type

Type TCustomUnit
    sCustomID       As String   ' �ŷ�ó ���� �ڵ�
    sArticleID      As String   ' ǰ������ �ڵ�
    sStuffWidthID   As String   ' ������
    sWorkID         As String   ' �������� �ڵ�
    nUnitPrice      As Single   ' �ܰ�
    nChunkRate      As Single   ' ����( Loss���� )
    sETC            As String   ' ���
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
    
    With adoCmd     '�ŷ�ó �Է�
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
        'S_201312_���_99 �� ���� �߰�-------------------------------------------------------
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, TItem.sOldNNewClss)       '���θ�,�����ּ� ���� 0:���θ�, 1:����
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sGunMoolMngNo)     '�ǹ����� �����ĺ���ȣ
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, TItem.sAddress1)      '���θ� �ּ�1
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, TItem.sAddress2)      '���θ� �ּ�2
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, TItem.sAddressAssist)    '���θ� ���� �ּ�
        '----------------------------------------------------------------------------------------
        'S_201312_���_99 �� ���� ����(OLD:tItem.sAddress1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sAddressJiBun1)
        'S_201312_���_99 �� ���� ����(OLD:tItem.sAddress2)
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
        'S_201312_���_99 �� ���� ����-.sOldNNewClss,.sGunMoolMngNo,.sAddress1,.sAddress2,.sAddressAssist �߰�
        '                              .sAddress1 -> sAddressJiBun1 �� ���� , .sAddress2 -> sAddressJiBun2 �� ����
        sLog(0) = "xp_Custom_iCustom '" & .sCustomID & "', '" & .sKCustom & "', '" & .sShortCustom & "', '" & .sECustom & "', '" & _
                .sCustomNo & "', '" & .sChief & "', '" & .sCondition & "', '" & .sCategory & "', '" & _
                .sZipCode & "', , '" & .sOldNNewClss & "', '" & .sGunMoolMngNo & "', '" & .sAddress1 & "', '" & .sAddress2 & "', '" & _
                .sAddressAssist & "', '" & .sAddressJiBun1 & "', '" & .sAddressJiBun2 & "',  '" & .sPhone1 & "', '" & _
                .sPhone2 & "', '" & .sFaxNO & "', '" & .sEMail & "', '" & .sHomePage & "', '" & _
                .sName & "', '" & .sPhone & "', '" & .sTradeID & "', '" & .sUserID & "', '" & _
                .sUserPassword & "', '" & .sLossClss & "', '" & .sSpendingClss & "', '" & .sWorkingClss & "', '" & _
                .sCalcClss & "', '" & .sPointClss & "'"
    End With
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewCustom = False
''        Err.Raise Err.Number, "CCustom.AddNewCustom", Err.Description
    Else   ' �α�
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
    
    With adoCmd     '�ŷ�ó ����
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
        'S_201312_���_99 �� ���� �߰�-------------------------------------------------------
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, TItem.sOldNNewClss)       '���θ�,�����ּ� ���� 0:���θ�, 1:����
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, TItem.sGunMoolMngNo)     '�ǹ����� �����ĺ���ȣ
        .Parameters.Append .CreateParameter(, adVarWChar, adParamInput, 80, TItem.sAddress1)      '���θ� �ּ�1
        .Parameters.Append .CreateParameter(, adVarWChar, adParamInput, 80, TItem.sAddress2)      '���θ� �ּ�2
        .Parameters.Append .CreateParameter(, adVarWChar, adParamInput, 80, TItem.sAddressAssist)    '���θ� ���� �ּ�
        '----------------------------------------------------------------------------------------
        'S_201312_���_99 �� ���� ����(OLD:tItem.sAddress1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, TItem.sAddressJiBun1)
        'S_201312_���_99 �� ���� ����(OLD:tItem.sAddress2)
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
        'S_201312_���_99 �� ���� ����-.sOldNNewClss,.sGunMoolMngNo,.sAddress1,.sAddress2,.sAddressAssist �߰�
        '                              .sAddress1 -> sAddressJiBun1 �� ���� , .sAddress2 -> sAddressJiBun2 �� ����
        sLog(0) = "xp_Custom_uCustom '" & .sCustomID & "', '" & .sKCustom & "', '" & .sShortCustom & "', '" & .sECustom & "', '" & _
                .sCustomNo & "', '" & .sChief & "', '" & .sCondition & "', '" & .sCategory & "', '" & _
                .sZipCode & "', '" & .sOldNNewClss & "', '" & .sGunMoolMngNo & "', '" & .sAddress1 & "', '" & .sAddress2 & "', '" & _
                .sAddressAssist & "', '" & .sAddressJiBun1 & "', '" & .sAddressJiBun2 & "',  '" & .sPhone1 & "', '" & _
                .sPhone2 & "', '" & .sFaxNO & "', '" & .sEMail & "', '" & .sHomePage & "', '" & _
                .sName & "', '" & .sPhone & "', '" & .sTradeID & "', '" & .sUserID & "', '" & _
                .sUserPassword & "', '" & .sLossClss & "', '" & .sSpendingClss & "', '" & .sWorkingClss & "', '" & _
                .sCalcClss & "', '" & .sPointClss & "'"
    End With
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdateCustom = False
''        Err.Raise Err.Number, "CCustom.UpdateCustom", Err.Description
    Else   ' �α�
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
    
    With adoCmd     '�ŷ�ó ����
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
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeleteCustom = False
''        Err.Raise Err.Number, "CCustom.DeleteCustom", Err.Description
    Else   ' �α�
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
    
    With adoCmd     '�ŷ�ó �ܰ� �Է�
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
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeleteCustomUnit = False
''        Err.Raise Err.Number, "CCustom.DeleteCustomUnit", Err.Description
    Else   ' �α�
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
    
    With adoCmd     '�ŷ�ó �ܰ� �Է�
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
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewCustomUnit = False
''        Err.Raise Err.Number, "CCustom.AddNewCustomUnit", Err.Description
    Else   ' �α�
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
'    With adoCmd     '�ŷ�ó �ܰ� �Է�
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
'    If bError Then  ' ���� �α�
'        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)
'
'        UpdateCustomUnit = False
'        Err.Raise Err.Number, "CCustom.UpdateCustomUnit", Err.Description
'    Else   ' �α�
'        Call LogService(sLog)
'
'        UpdateCustomUnit = True
'    End If


End Function

