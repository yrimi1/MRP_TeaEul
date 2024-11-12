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
'** System �� : PlusLib2
'** ����    : ������� DBó����
'** Author    : Wizard
'** �ۼ���    :
'** ����      :
'** ��������  :
'------------------------------------------------------------------------------------------------------------------
' * �����̷�
'------------------------------------------------------------------------------------------------------------------
' ����        �۾���  ��û��          ��û��ȣ           ��û���� �� ���泻��
'------------------------------------------------------------------------------------------------------------------
' 2013.12.12  ���¿�                 S_201312_���_99    �����ּҿ��� ���θ� �ּҷ� �Է°����ϰ�-��� ���� Type���� ����,AddNewPerson  �Լ�,UpdatePerson �Լ� ����
'*******************************************************************************

Option Explicit

Type TPerson
    sPersonID       As String    ' [0] �ڵ� (4)
    sUserID         As String    ' [1] Username
    sPassword       As String    ' [2] Password
    sName           As String    ' [3] ����
    sDepartID       As String    ' [5] �μ�
    sDutyID         As String    ' [6] ��å (0:����,1:����,2:��,3:�̻�,4:����,5:����,6:����,7:�븮,8:����,9:���������,10:������ ���)
    sTeamID         As String    ' �۾���
    sStartDate      As String    ' [7] �Ի�����
    sEndDate        As String    ' [8] �������
    sRegistID       As String    ' [9] �ֹε�Ϲ�ȣ
    sHandPhone      As String    ' [10] �ڵ���
    sPhone          As String    ' [11] ��ȭ��ȣ
    sBirthDay       As String    ' [12] �������
    sSolarClss      As String    ' [13] ��/���� ����
    sZipCode        As String    ' [14] ������ȣ
    'S_201312_���_99 �� ���� �߰�---------------------------------------------------
    sOldNNewClss        As String '[15]  ���θ�,�����ּ� ���� 0:���θ�, 1:����
    sGunMoolMngNo        As String '[16]  �ǹ����� �����ĺ���ȣ
    sAddress1       As String '[17] ���θ��ּ�1
    sAddress2       As String '[18] ���θ��ּ�2
    sAddressAssist       As String '[19] ���θ������ּ�
    '--------------------------------------------------------------------------------
    'S_201312_���_99 �� ���� ����(OLD:sAddress1)
    sAddressJiBun1       As String    ' [20] �� �ּ�1
    'S_201312_���_99 �� ���� ����(OLD:sAddress1)
    sAddressJiBun2       As String    ' [21] ���� �ּ�2
    sEMail          As String        ' [22] �̸���
    sRemark         As String    ' [23] �������
    sEName          As String    ' [24] �����̸�
    sSMSYN          As String   '  [25] SMS���� ���ۿ���(Y:���� ���� ���, N:��� �ƴ�)
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
        'S_201312_���_99 �� ���� �߰�-------------------------------------------------------
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, stData.sOldNNewClss)       '���θ�,�����ּ� ���� 0:���θ�, 1:����
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, stData.sGunMoolMngNo)     '�ǹ����� �����ĺ���ȣ
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddress1)      '���θ� �ּ�1
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddress2)      '���θ� �ּ�2
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddressAssist)    '���θ� ���� �ּ�
        '----------------------------------------------------------------------------------------
        'S_201312_���_99 �� ���� ����(OLD:stData.sAddress1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sAddressJiBun1)
        'S_201312_���_99 �� ���� ����(OLD:stData.sAddress2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, stData.sAddressJiBun2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sEMail)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sRemark)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, stData.sTeamID)               '�۾���
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, stData.sEName)           '�����̸�
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, stData.sSMSYN)                '�������� ��󿩺�
        
        .Execute
        
        PersonID = .Parameters(0).Value
    End With
    
    nSql = 1
    
    If nCnt <> 0 Then
        With adoCmd     ' ����ں� �޴� �Է�
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
        With adoCmd     ' ����ں� �޴� �Է�
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
        'S_201312_���_99 �� ���� ����-.sOldNNewClss,.sGunMoolMngNo,.sAddress1,.sAddress2,.sAddressAssist �߰�
        '                              .sAddress1 -> sAddressJiBun1 �� ���� , .sAddress2 -> sAddressJiBun2 �� ����
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

    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        AddNewPerson = False
''        Err.Raise Err.Number, "CPerson.AddNewPerson", Err.Description
    Else   ' �α�
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
        'S_201312_���_99 �� ���� �߰�-------------------------------------------------------
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, stData.sOldNNewClss)       '���θ�,�����ּ� ���� 0:���θ�, 1:����
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 25, stData.sGunMoolMngNo)     '�ǹ����� �����ĺ���ȣ
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddress1)      '���θ� �ּ�1
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddress2)      '���θ� �ּ�2
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, stData.sAddressAssist)    '���θ� ���� �ּ�
        '----------------------------------------------------------------------------------------
        'S_201312_���_99 �� ���� ����(OLD:stData.sAddress1)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sAddressJiBun1)
        'S_201312_���_99 �� ���� ����(OLD:stData.sAddress2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, stData.sAddressJiBun2)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sEMail)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, stData.sRemark)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, stData.sTeamID)              '�۾���
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, stData.sEName)           '�����̸�
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, stData.sSMSYN)                '�������� ��󿩺�

        .Execute
    End With
    
    nSql = 1
    
    With adoCmd     '����ں� �޴� ����
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Menu_dUserMenu"
    
        Call ClearParameter(adoCmd)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sPersonID)
        
        .Execute
    
    End With
    
    nSql = nSql + 1
    
    If nCnt <> 0 Then
        With adoCmd     ' ���ο� ����ں� �޴� �Է�
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
    
     With adoCmd     '����ں� �������� ���� ����
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_dPersonMachine"
    
        Call ClearParameter(adoCmd)
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, stData.sPersonID)
        
        .Execute
    
    End With
    
    nSql = nSql + 1
    
    If nMachineCnt <> 0 Then
        With adoCmd     ' ����ں� �޴� �Է�
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
        'S_201312_���_99 �� ���� ����-.sOldNNewClss,.sGunMoolMngNo,.sAddress1,.sAddress2,.sAddressAssist �߰�
        '                              .sAddress1 -> sAddressJiBun1 �� ���� , .sAddress2 -> sAddressJiBun2 �� ����
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

    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        UpdatePerson = False
''        Err.Raise Err.Number, "CPerson.UpdatePerson", Err.Description
    Else   ' �α�
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
    
    With adoCmd     '����ڻ���
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Person_dPerson"
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, sPersonID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, EndDate)
        
        
        .Execute
    
    End With
    
    With adoCmd     '����� �޴�����
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
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        DeletePerson = False
''        Err.Raise Err.Number, "CPerson.DeletePerson", Err.Description
    Else   ' �α�
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
    
    With adoCmd     '����ڻ���
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
    
    If bError Then  ' ���� �α�
        Call ErrLogService(sLog, Err.Number, Err.Description, nSql)

        ChangePassWord = False
''        Err.Raise Err.Number, "CPerson.UpdatePassWord", Err.Description
    Else   ' �α�
        Call LogService(sLog)

        ChangePassWord = True
    End If

End Function

' ���������� ���α׷����� ListBox(�����۾���) ä���
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

' ���������� ���α׷����� ListBox(�����۾���) ä���
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

