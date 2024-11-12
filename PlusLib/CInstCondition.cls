VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInstCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public g_adoCon As ADODB.Connection

Public Type TDelType
    xpProName      As String       '���ν��� ��
    OrderID        As String       '������ȣ
    Process        As String       '[] ������
    IndiDate       As String       '��������
    IndiTime       As String       '���ýð�
End Type


'--- �������� ( wi_Tenter ) ����ü
Public Type TTenter
    JobFlag        As String
    OrderID        As String
    Process        As String       '[] ������
    ProcessID      As String       '[] �����ڵ�
    IndiDate       As String       '��������
    IndiTime       As String       '���ýð�
    Temper         As Integer      '�µ�
    Velocity       As Integer      '�ӵ�
    OverFeed       As Double       'Over Feed
    Density        As Integer      '����е�
    SettingClss    As String       'Setting ����
    WorkCond       As String       '�۾�����
    CodeID         As String       '�ҷ������ڵ�
    DryID          As String       '��������
    PersonID       As String       '�ۼ����ڵ�
    Remark         As String       '���
End Type

'--- �������� ( wi_Dry )
Public Type TDry
    JobFlag        As String
    OrderID        As String
    Process        As String       '[] ������
    ProcessID      As String       '[] �����ڵ�
    IndiDate       As String       '��������
    IndiTime       As String       '���ýð�
    Temper         As Integer      '�µ�
    Velocity       As Integer      '�ӵ�
    OverFeed       As Double       'Over Feed
    CodeID         As String       '�ҷ������ڵ�
    DryID          As String       '��������
    PersonID       As String       '�ۼ����ڵ�
    Remark         As String       '���
End Type

'--- �������� ( wi_Refine )
Public Type TRefine
    JobFlag        As String
    OrderID        As String
    Process        As String       '[] ������
    ProcessID      As String       '[] �����ڵ�
    IndiDate       As String       '��������
    IndiTime       As String       '���ýð�
    Temper         As Integer      '�µ�
    Velocity       As Integer      '�ӵ�
    RefineClss     As String       '���ñ���
    SettingClss    As String      'Setting����
    PersonID       As String       '�ۼ����ڵ�
    Remark         As String       '���
End Type


'--- CPB ��ó�� ����( wi_CPBPre )
Public Type TCPBPre
    JobFlag        As String
    OrderID        As String
    Process        As String       '[] ������
    ProcessID      As String       '[] �����ڵ�
    IndiDate       As String       '��������
    IndiTime       As String       '���ýð�
    Velocity       As Integer      '�ӵ�
    RefineClss     As String       '���ñ���
    PersonID       As String       '�ۼ����ڵ�
    Remark         As String       '���
End Type


'--- Peach ����( wi_peach )
Public Type TPeach
    JobFlag        As String
    OrderID        As String
    Process        As String       '[] ������
    ProcessID      As String       '[] �����ڵ�
    IndiDate       As String       '[] ��������
    IndiTime       As String       '[] ���ýð�
    Velocity       As Integer      '[] �ӵ�
    PePaBon1       As Integer      '[] ���ĺ�1
    PePaBon2       As Integer      '[] ���ĺ�2
    PePaBon3       As Integer      '[] ���ĺ�3
    PePaBon4       As Integer      '[] ���ĺ�4
    Density        As Integer      '[] �е�
    Tention        As Integer      '[] �е�
    Pressure1      As Integer      '[] ���ĺ�1
    Pressure2      As Integer      '[] ���ĺ�2'
    Pressure3      As Integer      '[] ���ĺ�3
    PersonID       As String       '[] �ۼ����ڵ�
    Remark         As String       '[] ���
End Type


'--- Peach ����( wi_Mercerizng )
Public Type TMercerizing
    OrderID        As String
    Process        As String       '[] ������
    ProcessID      As String       '[] �����ڵ�
    IndiDate       As String       '[] ��������
    IndiTime       As String       '[] ���ýð�
    Rpm            As Integer      '[] rpm
    Temper         As Integer      '[] �µ�
    DyeAuxID       As Integer      '[] �����ڵ�
    Density        As Integer      '[] ������
    PersonID       As String       '[] �ۼ����ڵ�
    Remark         As String       '[] ���
End Type

Public g_sUserName As String
Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property
Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

'******************************************************************************
' ��   ��  �� : GetInstRecord
' ��       �� : �۾��������� ���ڵ� �о���� ( wi_CPBPre, wi_Refine, wi_Tenter, wi_Peach, wi_Dry )
' �� �� �� �� : a = GetInstRecord( arg1, arg2, arg3, arg4, arg5 )
' �� �� �� �� : arg1 -> Stored procedure name( Table�� ���� �ٸ� )
'             : arg2 -> ������ȣ
'             : arg3 -> ������ȣ�� ã��(1) / ���Ͼ�ü + ǰ������ ã��(2)
'             : arg4 -> ������(1) / ����(0)
'             : arg5 -> ������
' ��   ��  �� : a = �ǵ��� �޴� Recordset
'******************************************************************************

Public Function GetInstRecord(ByVal sXpName As String _
                            , ByVal sOrderID As String _
                            , ByVal nChkOrderID As Integer _
                            , ByVal nChkProcessID As Integer _
                            , ByVal sProcessName As String _
                            , ByVal nRows As Integer) As ADODB.Recordset
                            
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = sXpName

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkProcessID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sProcessName)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nRows)

        Set GetInstRecord = .Execute
    End With
    Set adoCmd = Nothing

End Function
Public Function GetInstOneRec(ByVal sXpName As String _
                            , ByVal sOrderID As String _
                            , ByVal sProcName As String _
                            , ByVal sIndiDate As String _
                            , ByVal sIndiTime As String) As ADODB.Recordset
                            
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_instCondi_OneRec"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, sXpName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sOrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, sProcName)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sIndiDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sIndiTime)

        Set GetInstOneRec = .Execute
    End With
    Set adoCmd = Nothing

End Function
Public Function GetInstDefectList(ByVal sProcessID As String) As ADODB.Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_InstCondi_sDefect"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sProcessID)

        Set GetInstDefectList = .Execute
    End With
    Set adoCmd = Nothing

End Function

''Public Function aaaaa() As Recordset
''    Dim adoCmd As ADODB.Command
''
''    Set adoCmd = New ADODB.Command
''
''    With adoCmd
''        .ActiveConnection = g_adoCon
''        .CommandType = adCmdStoredProc
''        .CommandText = "xp_ProcessWait_sOrder"
''
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, "")
''
''    End With
''    Set aaaaa = adoCmd.Execute
''    Set adoCmd = Nothing
''End Function


' ����CPB�� ������ �ű�/����
Public Function AddNewCPBPre(NewItem As TCPBPre) As Boolean
    Dim adoCmd As ADODB.Command
    Dim AffectedRows_int As Integer
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_InstCondi_iuCPBPre"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, AffectedRows_int)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.JobFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, NewItem.OrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, NewItem.Process)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.IndiDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.IndiTime)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 4, NewItem.Velocity)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 6, NewItem.RefineClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.PersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, NewItem.Remark)
        
        .Execute
        AffectedRows_int = .Parameters(0).Value
    End With
    Set adoCmd = Nothing
    
    If AffectedRows_int = 1 Then
        AddNewCPBPre = True
    Else
        AddNewCPBPre = False
    End If
    
End Function

' ����CPB�� ������ �ű�/����
Public Function AddNewRefine(NewItem As TRefine) As Boolean
    Dim adoCmd As ADODB.Command
    Dim AffectedRows_int As Integer
    
    
''xp_InstCondi_iuRefine
''    @nAffected      AS  smallint  OUTPUT
''  , @sJobFlag       AS  char(1)
''  , @sOrderId       AS  char(10)
''  , @sProcess       AS  nvarchar(20)
''  , @sIndiDate      AS  char(8)
''  , @sIndiTime      AS  char(4)
''  , @nTemper        AS  smallint        -- �µ�
''  , @nVelocity      AS  smallint        -- �ӵ�
''  , @sRefineClss    AS  nvarchar(6)     -- ���ñ���
''  , @sSettingClss   AS  nvarchar(10)    -- Setting ����
''  , @sPersonID      AS  char(8)         -- �ۼ��� �ڵ�
''  , @sRemark        AS  nvarchar(50)    -- ���
    
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_InstCondi_iuRefine"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, AffectedRows_int)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.JobFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, NewItem.OrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, NewItem.Process)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.IndiDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.IndiTime)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Temper)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Velocity)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 6, NewItem.RefineClss)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, NewItem.SettingClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.PersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, NewItem.Remark)
        
        .Execute
        AffectedRows_int = .Parameters(0).Value
    End With
    Set adoCmd = Nothing
    
    If AffectedRows_int = 1 Then
        AddNewRefine = True
    Else
        AddNewRefine = False
    End If
End Function

'--- Tenter ���ڵ� �ű�/���� ���μ���
Public Function AddNewTenter(NewItem As TTenter) As Boolean
    Dim adoCmd As ADODB.Command
    Dim AffectedRows_int As Integer
    
''    xp_InstCondi_iuTenter
''    @nAffected      AS  smallint  OUTPUT
''  , @sJobFlag       AS  char(1)
''  , @sOrderId       AS  char(10)
''  , @sProcess       AS  nvarchar(20)
''  , @sIndiDate      AS  char(8)
''  , @sIndiTime      AS  char(4)
''  , @nTemper        AS  smallint
''  , @nVelocity      AS  smallint
''  , @nOverFeed      AS  numeric(3, 1)
''  , @nDensity       AS  smallint
''  , @sSettingClss   AS  nvarchar(10)
''  , @sWorkCond      AS  char(2)
''  , @sCodeID        AS  char(3)   -- �ҷ������ڵ�
''  , @sDryID         AS  char(2)
''  , @sPersonID      AS  char(8)
''  , @sRemark        AS  nvarchar(50)
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_InstCondi_iuTenter"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, AffectedRows_int)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.JobFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, NewItem.OrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, NewItem.Process)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.IndiDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.IndiTime)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Temper)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Velocity)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 1, NewItem.OverFeed)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Density)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, NewItem.SettingClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewItem.WorkCond)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewItem.CodeID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, NewItem.DryID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.PersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, NewItem.Remark)
        
        .Execute
       AffectedRows_int = .Parameters(0).Value
    End With
    Set adoCmd = Nothing
    
    If AffectedRows_int = 1 Then
        AddNewTenter = True
    Else
        AddNewTenter = False
    End If

End Function


'--- Peach  ���ڵ� �ű�/���� ���μ���
Public Function AddNewPeach(NewItem As TPeach) As Boolean
    Dim adoCmd As ADODB.Command
    Dim AffectedRows_int As Integer
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_InstCondi_iuPeach"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, AffectedRows_int)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.JobFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, NewItem.OrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, NewItem.Process)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.IndiDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.IndiTime)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Velocity)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.PePaBon1)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.PePaBon2)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.PePaBon3)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.PePaBon4)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Density)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Tention)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Pressure1)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Pressure2)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Pressure3)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.PersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, NewItem.Remark)
        
        .Execute
        AffectedRows_int = .Parameters(0).Value
        
    End With
    Set adoCmd = Nothing
    
    If AffectedRows_int = 1 Then
        AddNewPeach = True
    Else
        AddNewPeach = False
    End If

End Function


Public Function AddNewDry(NewItem As TDry) As Boolean
    Dim adoCmd As ADODB.Command
    Dim AffectedRows_int As Integer
    
''xp_InstCondi_iuDry
''    @nAffected      AS  smallint  OUTPUT
''  , @sJobFlag       AS  char(1)
''  , @sOrderId       AS  char(10)
''  , @sProcess       AS  nvarchar(20)
''  , @sIndiDate      AS  char(8)
''  , @sIndiTime      AS  char(4)
''  , @nTemper        AS  smallint        -- �µ�
''  , @nVelocity      AS  smallint        -- �ӵ�
''  , @nOverFeed      AS  numeric(3, 1)   -- Over Feed
''  , @sCodeID        AS  char(3)         -- �ҷ������ڵ�
''  , @sDryID         AS  char(2)         -- ��������
''  , @sPersonID      AS  char(8)         -- �ۼ��� �ڵ�
''  , @sRemark        AS  nvarchar(50)    -- ���
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_InstCondi_iuDry"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, AffectedRows_int)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, NewItem.JobFlag)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, NewItem.OrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, NewItem.Process)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.IndiDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, NewItem.IndiTime)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Temper)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, NewItem.Velocity)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 1, NewItem.OverFeed)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, NewItem.CodeID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 3, NewItem.DryID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, NewItem.PersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, NewItem.Remark)
        
        .Execute
        AffectedRows_int = .Parameters(0).Value
    End With
    Set adoCmd = Nothing
    
    If AffectedRows_int = 1 Then
        AddNewDry = True
    Else
        AddNewDry = False
    End If
    
End Function


' ��� ������ ���� �������� ���
' ��� ������ ���� �������� ���
Public Function DelInstCondition(DelItem As TDelType) As Boolean
    Dim adoCmd As ADODB.Command
    Dim AffectedRows_int As Integer
    
    Set adoCmd = New ADODB.Command
    
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_InstCondi_Delete"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, AffectedRows_int)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, DelItem.OrderID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, DelItem.Process)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, DelItem.IndiDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, DelItem.IndiTime)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, DelItem.xpProName)
        .Execute
        AffectedRows_int = .Parameters(0).Value
    End With
    Set adoCmd = Nothing
    
    If AffectedRows_int = 1 Then
        DelInstCondition = True
    Else
        DelInstCondition = False
    End If

End Function


Public Function GetDraftOrder(bChkDate As Boolean, sFromDate As String, sToDate As String, _
                            Optional sCustom As String, Optional sOrder As String, _
                            Optional nFlag As Integer) As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    Dim nChkFlag(2) As Integer
    
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Order_sDraftOrder"
        
        nChkFlag(0) = IIf(bChkDate, 0, 1)           '[1] �������� ����
        nChkFlag(1) = IIf(Len(sCustom) > 0, 1, 0)   '[2] �ŷ�ó ����
        If Len(sOrder) = 0 Then
            nChkFlag(2) = 0
        Else
            nChkFlag(2) = nFlag + 1                        '[3] ������ȣ, Order No. ����
        End If
        
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkFlag(0))
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sFromDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sToDate)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkFlag(1))
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustom)
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, 0)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 15, "")
        .Parameters.Append .CreateParameter(, adTinyInt, adParamInput, 1, nChkFlag(2))
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 15, sOrder)
        
        Set GetDraftOrder = .Execute
    End With

    Set adoCmd = Nothing
End Function

'''����CPB�� �����ͻ���
''Public Function DelCPBPre(DelItem As TCPBPre) As Boolean
''    Dim adoCmd As ADODB.Command
''    Dim AffectedRows_int As Integer
''
''    Set adoCmd = New ADODB.Command
''
''    With adoCmd
''        .ActiveConnection = g_adoCon
''        .CommandType = adCmdStoredProc
''        .CommandText = "xp_InstCondi_dPBPre"
''
''        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 1, AffectedRows_int)
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, DelItem.OrderId)
''        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, DelItem.ProcessID)
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, DelItem.IndiDate)
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, DelItem.IndiTime)
''        .Execute
''    End With
''    Set adoCmd = Nothing
''
''    If AffectedRows_int = 1 Then
''        DelCPBPre = True
''    Else
''        DelCPBPre = False
''    End If
''
''End Function