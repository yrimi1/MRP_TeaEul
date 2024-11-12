VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRapid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Type tRapidCard
    sCardID     As String
    sSplitID    As String
    lDyeSchID   As Long
End Type

Private Const CLASS_NAME As String = "CRapid"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(NewUser As String)
    g_sUserName = NewUser
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

' ������⿡ �ִ� ī�� ���� ����Ʈ
Public Function GetRapidScheduling(nChkValue As Integer, nChkSchID As Long) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyeScheduleList"

        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkValue)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nChkSchID)
        
        Set GetRapidScheduling = .Execute
    End With

    Set adoCmd = Nothing
End Function

' ������⿡ �ִ� ī����� Box, Soaping ����Ʈ
Public Function GetRapidSchedulingBox(nChkSchID As Long, nChkSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyeScheduleListBox"

        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nChkSchID)
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nChkSeq)
        
        Set GetRapidSchedulingBox = .Execute
    End With

    Set adoCmd = Nothing
End Function

' �������ð� ������ ����Ʈ
Public Function GetRapidScheduledData() As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyeScheduledData"

        Set GetRapidScheduledData = .Execute
    End With

    Set adoCmd = Nothing
End Function

' �� ȣ�⺰ �ֱ� ���������� �ѰǾ� ��������
Public Function GetRapidWorkedEachData() As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyeWorkedTwoByMachine"

        Set GetRapidWorkedEachData = .Execute
    End With

    Set adoCmd = Nothing
End Function

' �ش� ����ȣ���� ù��° ��������� �׸� ��������
'Public Function GetCheckFirstWork(sProcID As String, iMachID As Integer) As ADODB.Recordset
'    Dim adoCmd As ADODB.Command
'
'    Set adoCmd = New ADODB.Command
'    With adoCmd
'        .ActiveConnection = g_adoCon
'        .CommandType = adCmdStoredProc
'        .CommandText = "xp_Rapid_sCheckFirstWork"
'
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcID)
'        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, iMachID)
'
'        Set GetCheckFirstWork = .Execute
'
'    End With
'    Set adoCmd = Nothing
'End Function
' �ش� ����ȣ���� ù��° ��������� �׸� ��������
Public Function GetCheckFirstWork(sProcID As String, iMachID As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sCheckFirstWork"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, iMachID)
        
        Set GetCheckFirstWork = .Execute
                
    End With
    Set adoCmd = Nothing
End Function
' �������ó����� �� �� �۾������� üũ
Public Function GetCheckDyeWorking(nSchID As Long, nSchSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyeWorking"
        
        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nSchID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nSchSeq)
        
        Set GetCheckDyeWorking = .Execute
    End With

    Set adoCmd = Nothing
End Function

' ��������(�۾��� �̿Ϸ��)�� ������ ī������ üũ
Public Function GetCheckDyeSch(sCardID As String, sSplitID As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sCheckSchedule"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sSplitID)
        
        Set GetCheckDyeSch = .Execute
    End With

    Set adoCmd = Nothing
End Function

' �� ���� ���� �� ����� ��Ȳ
Public Function GetResultProdDyeTrouble(sYear As String, sMonth As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Result_sProdDyeTrouble"
        
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sYear)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 2, sMonth)
        
        Set GetResultProdDyeTrouble = .Execute
    End With

    Set adoCmd = Nothing
End Function

' ���������� ���α׷����� ListBox(����ȣ��) ä���
Public Function GetMachineNoList(sChkValue As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sMachineNoList"

        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 20, sChkValue)
        
        Set GetMachineNoList = .Execute
    End With

    Set adoCmd = Nothing
End Function


Public Function GetMachineNo(sProcessID As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sMachineNo"

        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 4, sProcessID)
        
        Set GetMachineNo = .Execute
    End With

    Set adoCmd = Nothing
End Function


' ���������� ���α׷����� ListBox(��������) ä���
Public Function GetDyePatternList(nDyeKind As Integer, nDyeID As Integer, nPTNo As Integer) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sDyePatternGroup"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeKind)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nPTNo)
        
        Set GetDyePatternList = .Execute
    End With

    Set adoCmd = Nothing
End Function

' �����۾� ���� �Է�
Public Function AddNewwiRapidItem(CardList() As tRapidCard, lDyeSchID As Long, sProcID As String, sMachID As String, _
                                iRapidSeq As Integer, sWorkClss As String, sRapidClss As String, _
                                sPatternID As String, iUnitWght As Integer, lRoll As Long, lQty As Long, _
                       sInstClss As String, sPersonID As String, sRemark As String, Optional nDyeSchID As Long = 0, Optional nDyeSeq As Integer = 0) As Boolean
    Dim adoCmd As ADODB.Command
    Dim iDyeSchID As Long
    Dim iDyeSeq As Integer
    Dim i%, j%, K&

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        
        .CommandText = "xp_Rapid_iwiRapid"
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInputOutput, 4, lDyeSchID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, 0)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sMachID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, iRapidSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sWorkClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sRapidClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sPatternID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, iUnitWght)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, lRoll)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, lQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sInstClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, sRemark)
        .Execute
        
        iDyeSchID = .Parameters(0).Value
        iDyeSeq = .Parameters(1).Value
        nDyeSchID = .Parameters(0).Value
        nDyeSeq = .Parameters(1).Value
        
        .CommandText = "xp_Rapid_iwiRapidSub"
        .Prepared = True

        For j = 0 To UBound(CardList)
            Call ClearParameter(adoCmd)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, iDyeSchID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, iDyeSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, j + 1)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, CardList(j).sCardID)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, CardList(j).sSplitID)
            .Execute
        Next j
    End With
    
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    AddNewwiRapidItem = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "���嵵�� ����"
    AddNewwiRapidItem = False
End Function

Public Function DeletewiRapid(nDyeSchID As Long, nDyeSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_dwiRapid"

        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nDyeSchID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
        .Execute
    End With

    g_adoCon.CommitTrans
    DeletewiRapid = True
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    DeletewiRapid = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "�������� ����"
End Function

Public Function DeletewkRapid(nDyeSchID As Long, nDyeSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_dwkRapidwkResult"

        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nDyeSchID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
        .Execute
    End With

    g_adoCon.CommitTrans
    DeletewkRapid = True
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    DeletewkRapid = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "�������� ����"
End Function

Public Function UpdateRapidSeq(nDyeSchID As Long, nDyeSeq As Integer, sWaitProcID As String, _
                                sMachID As String, nRapidSeq As Integer, nPlusMinus As Integer) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_uRapidSeq"

        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nDyeSchID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sWaitProcID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sMachID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nRapidSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nPlusMinus)
        .Execute
    End With

    UpdateRapidSeq = True
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    UpdateRapidSeq = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "������Ʈ ����"
End Function

' �����۾� ���� �Է�
'Public Function UpdatewiRapid(nDyeSchID As Long, nDyeSeq As Integer, sProcID As String, sMachID As String, _
'                                nRapidSeq As Integer, sWorkClss As String, sRapidClss As String, _
'                                sPatternID As String, nUnitWght As Integer, nRoll As Long, nQty As Long, _
'                                sPersonID As String, sRemark As String, nChkChange As Integer) As Boolean
'    Dim adoCmd As ADODB.Command
'
'    On Error GoTo ErrHandler
'
'    Set adoCmd = New ADODB.Command
'    With adoCmd
'        .ActiveConnection = g_adoCon
'        .CommandType = adCmdStoredProc
'        .CommandText = "xp_Rapid_uwiRapidItem"
'
'        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nDyeSchID)
'        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcID)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sMachID)
'        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nRapidSeq)
'        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sWorkClss)
'        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sRapidClss)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sPatternID)
'        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nUnitWght)
'        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nRoll)
'        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nQty)
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPersonID)
'        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, sRemark)
'        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkChange)
'        .Execute
'    End With
'
'    Set adoCmd = Nothing
'    UpdatewiRapid = True
'
'    Exit Function
'
'ErrHandler:
'    g_adoCon.RollbackTrans
'    Set adoCmd = Nothing
'    UpdatewiRapid = False
'    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "������Ʈ ����"
'End Function

' �����۾� ���� �Է�
Public Function UpdatewiRapid(nDyeSchID As Long, nDyeSeq As Integer, sProcID As String, sMachID As String, _
                                nRapidSeq As Integer, sWorkClss As String, sRapidClss As String, _
                                sPatternID As String, nUnitWght As Integer, nRoll As Long, nQty As Long, _
                                sPersonID As String, sRemark As String, nChkChange As Integer) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_uwiRapidItem"
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nDyeSchID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sProcID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sMachID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nRapidSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sWorkClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sRapidClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sPatternID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nUnitWght)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nRoll)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, sRemark)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkChange)
        .Execute
    End With
    
    Set adoCmd = Nothing
    UpdatewiRapid = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    UpdatewiRapid = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "������Ʈ ����"
End Function

Public Function GetWaitWorkDyeProcCard(nDyeSchID As Long, nDyeSeq As Integer, sWaitWork As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sWaitWorkDyeProc"
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 2, nDyeSchID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sWaitWork)
        
        Set GetWaitWorkDyeProcCard = .Execute
    End With
    Set adoCmd = Nothing
End Function

Public Function GetwiRapidData(nDyeSchID As Long, nDyeSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_swiRapid"
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 2, nDyeSchID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
        
        Set GetwiRapidData = .Execute
    End With
    Set adoCmd = Nothing
End Function

Public Function GetOrderColorList(sCardID As String, sSplitID As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Card_sOrderColorList"
        
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sSplitID)
        
        Set GetOrderColorList = .Execute
    End With
    Set adoCmd = Nothing
End Function

' �����۾� ���۽� ��������(wk_Rapid), ����������(wk_RapidSub), ��������(wk_Result) �ۼ�
'                 ī��(Card)������ UseClss = '�۾�' ���� ������Ʈ
Public Function AddwkRapid(nDyeSchID As Long, nDyeSeq As Integer, sWorkJo As String, sClss As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim iChkDye As Integer

    On Error GoTo ErrHandler

    If sClss = "����" Then
        iChkDye = 0
    Else
        iChkDye = 1
    End If
    
    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_iwkRapidwkResult"
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nDyeSchID)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nDyeSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sWorkJo)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, iChkDye)
        .Execute
    End With
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    AddwkRapid = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "���嵵�� ����"
    AddwkRapid = False
End Function

' �����۾����� �ۼ��Ϸ�� ��������(wk_Rapid), ��������(wk_Result), ī��(Card), ���ϰ�ȹ(pl_Pattern) ������Ʈ
Public Function UpdatewkRapid(nDyeSchID As Long, nDyeSeq As Integer, sPatternID As String, sWorkClss As String, _
                            sRapidClss As String, sPersonID As String, sWorkJo As String, sRemark As String, _
                            sEndDate As String, sEndTime As String, sClss As String, nChkHold As Integer, sHoldReason As String) As Boolean
    Dim adoCmd As ADODB.Command
    Dim iChkDye As Integer

    On Error GoTo ErrHandler

    If sClss = "����" Then
        iChkDye = 0
    Else
        iChkDye = 1
    End If

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_uwkRapidwkResult"
    
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nDyeSchID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sPatternID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sWorkClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sRapidClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPersonID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sWorkJo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, sRemark)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEndDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sEndTime)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, iChkDye)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkHold)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, sHoldReason)
        
        .Execute
    End With
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    UpdatewkRapid = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "���嵵�� ����"
    UpdatewkRapid = False
End Function

' �����߰��۾���
Public Function AddDyeWorkRapid(nDyeSchID As Long, nDyeSeq As Integer, sEndDate As String, sEndTime As String, nNewDyeSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_iAddDyeWork"
        
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nDyeSchID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEndDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sEndTime)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, nNewDyeSeq)
        
        .Execute
        
        nNewDyeSeq = CheckNum(.Parameters(4).Value)
    End With
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    AddDyeWorkRapid = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "���嵵�� ����"
    AddDyeWorkRapid = False
End Function

Public Function UpdateCardColor(sCardID As String, sSplitID As String, nOrderSeq As Integer, sPersonID As String) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Card_uColorChange"
    
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sSplitID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sPersonID)
        .Execute
    End With
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    UpdateCardColor = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "���嵵�� ����"
    UpdateCardColor = False
End Function

' �������� ���� SELECT
Public Function GetDyeWorkProcess() As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Work_sDyeProcess"
        
        Set GetDyeWorkProcess = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

' ������� ���� SELECT
Public Function GetDyeMachine(Optional NewKey As String) As Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Process_sMachine"
        If Len(NewKey) > 0 Then
            .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, NewKey)
        End If
        Set GetDyeMachine = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

' ���� ����� �޾ƿ���
Public Function GetDyePlant() As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Work_sDyeMachine"

        Set GetDyePlant = .Execute
    End With
    Set adoCmd = Nothing
End Function

' ���� ���� ȣ��
Public Function GetDyeMachineByPlant(sMachine As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Work_sDyeMachineNO"
        
        ' ���� �̸����� �˻�
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sMachine)
        
        Set GetDyeMachineByPlant = .Execute
                
    End With
    Set adoCmd = Nothing
End Function

' �������� ����
Public Function GetRapidResultByPlant(sSDate As String, sEDate As String, nChkProcessID As Integer, _
                                    sProcess As String, nChkMachineID As Integer, sMachineID As String, _
                                    nChkTeamID As Integer, sTeamID As String, nChkOrderID As Integer, _
                                    sOrderID As String, nChkCustomID As Integer, sCustomID As String, _
                                    nChkArticleID As Integer, sArticleID As String, nPlan As Integer, nResult As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Work_sResultWithRapid"
                
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkProcessID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sProcess)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkMachineID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 3, Format(sMachineID, "00"))
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkTeamID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sTeamID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(.CommandText, adTinyInt, adParamInput, 1, nPlan)
        .Parameters.Append .CreateParameter(.CommandText, adTinyInt, adParamInput, 1, nResult)
        
        Set GetRapidResultByPlant = .Execute
    End With
    Set adoCmd = Nothing

End Function


' ���������� ���� ����
Public Function GetRapidScheduleUse(sProgID As String) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_sScheduleUse"
        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 3, sProgID)

        Set GetRapidScheduleUse = .Execute
    End With

    Set adoCmd = Nothing
End Function


' ���������� ���� ���� ������Ʈ
Public Function UpdateScheduleUse(sProgID As String, sUseClss As String) As Boolean
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Rapid_uScheduleUse"
    
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, sProgID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sUseClss)
        
        .Execute
    End With
    g_adoCon.CommitTrans
    
    Set adoCmd = Nothing
    UpdateScheduleUse = True
    
    Exit Function

ErrHandler:
    g_adoCon.RollbackTrans
    Set adoCmd = Nothing
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "���嵵�� ����"
    UpdateScheduleUse = False
End Function
