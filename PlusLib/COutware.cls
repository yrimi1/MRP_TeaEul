VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COutWare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Type TOUTWARE
    OrderID     As String
    OutSeq      As Integer
    OutClss     As String
    WorkID      As String
    ExchRate    As Single
    UnitPrice   As Single
    OutCustom   As String
    LossRate    As Single
    LossQty     As Single
    OutRoll     As Integer
    OutQty      As Long
    OutRealQty  As Single
    OutDate     As String
    ResultDate  As String
    OutTime     As String
    BoOutClss   As String
    BoConfirmClss   As String
    BoConfirmDate   As String
    LoadTime    As String
    TranNo      As String
    TranSeq     As Integer
    OutType     As String
    Remark      As String
End Type

Type TOUTWARESUB
    OrderID   As String
    OutSeq    As Integer
    OutSubSeq   As Integer
    OrderSeq    As Integer
    RollSeq     As Integer
    BoxNo       As Integer
    LotNo       As String
    OutQty      As Integer
End Type

Private Const CLASS_NAME As String = "COutWare"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function GetOrderList(nChkDate As Integer, sDate As String, eDate As String, _
         nChkCustomID As Integer, sCustomID As String, _
         nChkArticleID As Integer, sArticleID As String, _
         nChkOrder As Integer, sOrder As String) As Recordset
    
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOrderList"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)

        Set GetOrderList = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetOrderSubTotal(sOrderID As String, Optional nChkOrderSeq As Integer = 0, Optional nOrderSeq As Integer = 0) As Recordset
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command

    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOrderSubTotal"

        .Parameters.Append .CreateParameter(.CommandText, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nChkOrderSeq)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nOrderSeq)

        Set GetOrderSubTotal = .Execute
    End With

    Set adoCmd = Nothing
End Function

Public Function GetOutware(nChkDate As Integer, sSDate As String, sEDate As String, _
            nChkCustomID As Integer, sCustomID As String, _
            nChkArticleID As Integer, sArticleID As String, _
            nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOrder"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)

        Set GetOutware = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutware"), Err.Description
End Function


Public Function GetOutwareSub(sOrderID As String, nOutSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOutwareSub"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOutSeq)

        Set GetOutwareSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutwareSub"), Err.Description
End Function

Public Function GetOutwareSubTotal(sOrderID As String, nOutSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOutwareSubTotal"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOutSeq)

        Set GetOutwareSubTotal = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutwareSubTotal"), Err.Description
End Function

Public Function GetOutwareSumByColor(sOrderID As String, nOutSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOutwareSumByColor"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOutSeq)

        Set GetOutwareSumByColor = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutwareSumByColor"), Err.Description
End Function


Public Function GetOrderOne(sOrderID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOrderOne"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)

        Set GetOrderOne = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOrderOne"), Err.Description
End Function

Public Function GetOutwareOn(nChkDate As Integer, sSDate As String, sEDate As String, _
            nChkCustomID As Integer, sCustomID As String, _
            nChkArticleID As Integer, sArticleID As String, _
            nChkOrder As Integer, sOrder As String, nBoOutClss As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOutwareOn"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nBoOutClss)

        Set GetOutwareOn = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutwareOn"), Err.Description
End Function

Public Function GetOutwareResult(nChkDate As Integer, sSDate As String, sEDate As String, _
            nChkCustomID As Integer, sCustomID As String, _
            nChkArticleID As Integer, sArticleID As String, _
            nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOutwareResult"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)

        Set GetOutwareResult = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutwareResult"), Err.Description
End Function

Public Function GetOutwareDetail(nChkDate As Integer, sSDate As String, sEDate As String, _
                                nChkCustomID As Integer, sCustomID As String, _
                                nChkArticleID As Integer, sArticleID As String, _
                                nChkOrder As Integer, sOrder As String, _
                                nChkOutClss As Integer, nOutClss As Integer, _
                                nTaxClss As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOutwareDetail"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOutClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOutClss)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nTaxClss)
        
        Set GetOutwareDetail = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutwareDetail"), Err.Description
End Function


Public Function GetOutwareOrder(ByVal nChkDate As Integer, ByVal sDate As String, ByVal eDate As String _
             , ByVal nChkCustomID As Integer, ByVal sCustomID As String _
             , ByVal nChkArticleID As Integer, ByVal sArticleID As String _
             , ByVal nChkOrder As Integer, ByVal sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOutwareOrder"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)

        Set GetOutwareOrder = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutwareOrder"), Err.Description
End Function

Public Function GetOutwareOrderDetail(sOrderID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sOutwareOrderDetail"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)

        Set GetOutwareOrderDetail = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutwareOrderDetail"), Err.Description
End Function

Public Function GetInspect(sOrderID As String, nChkDate As Integer, sSDate As String, sEDate As String, _
            nChkGradeID As Integer, sGradeID As String, _
            nChkColor As Integer, sColor As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sInspect"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkGradeID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sGradeID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkColor)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, sColor)

        Set GetInspect = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInspect"), Err.Description
End Function

Public Function GetInspectByColorSum(sOrderID As String, nChkDate As Integer, sSDate As String, sEDate As String, _
            nChkGradeID As Integer, sGradeID As String, _
            nChkColor As Integer, sColor As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_sInspectByColorSum"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkGradeID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sGradeID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkColor)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, sColor)

        Set GetInspectByColorSum = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetInspectByColorSum"), Err.Description
End Function


Public Function AddNewOutwareHT(ow As TOUTWARE, owSub() As TOUTWARESUB) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Insert
        .CommandText = "xp_Outware_iOutware"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, ow.OrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, ow.OutSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow.OutClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow.WorkID)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.ExchRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.UnitPrice)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, ow.OutCustom)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.LossRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.LossQty)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, ow.OutRoll)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, ow.OutQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 8, ow.OutRealQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, ow.OutDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, ow.ResultDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow.BoOutClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow.OutTime)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow.LoadTime)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow.OutType)

        .Execute
        iSQL = iSQL + 1

        ow.OutSeq = .Parameters(1).Value
        
        ' Insert Sub
        .CommandText = "xp_Outware_iOutwareSub"
        .Prepared = True

        For i = 0 To UBound(owSub)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, owSub(i).OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, ow.OutSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).OutSubSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).OrderSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).RollSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).BoxNo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, owSub(i).LotNo)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, owSub(i).OutQty)

            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COutWare.AddNewOutwareHT"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(owSub) + 1)
    With ow
        sLog(0) = "xp_Outware_iOutware '" & .OrderID & "', " & .OutSeq & ", '" & .OutClss & "', '" & _
            .WorkID & "', " & .ExchRate & ", " & .UnitPrice & ", '" & .OutCustom & "', " & .LossRate & ", " & _
            .LossQty & ", " & .OutRoll & ", " & .OutQty & ", " & .OutRealQty & ", '" & .OutDate & "', '" & _
            .ResultDate & "', '" & .BoOutClss & "', '" & .OutTime & "', '" & .LoadTime & "', '" & .OutType & "' "
    End With
    For i = 0 To UBound(owSub)
        With owSub(i)
            sLog(i + 1) = "xp_Outware_iOutwareSub '" & .OrderID & "', " & .OutSeq & ", " & .OutSubSeq & ", " & _
            .OrderSeq & ", " & .RollSeq & ", " & .BoxNo & ", '" & .LotNo & "', " & .OutQty & " "
        End With
    Next i
        
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        AddNewOutwareHT = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewOutwareHT"), Err.Description
    Else
        Call LogService(sLog)

        AddNewOutwareHT = True
    End If
End Function

Public Function AddNewOutware(ow As TOUTWARE, owSub() As TOUTWARESUB) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, j%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        ' Insert
        .CommandText = "xp_Outware_iOutware"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, ow.OrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, ow.OutSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow.OutClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow.WorkID)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.ExchRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.UnitPrice)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, ow.OutCustom)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.LossRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.LossQty)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, ow.OutRoll)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, ow.OutQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 8, ow.OutRealQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, ow.OutDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, ow.ResultDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow.BoOutClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow.OutTime)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow.LoadTime)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow.OutType)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, ow.Remark)

        .Execute
        iSQL = iSQL + 1

        ow.OutSeq = .Parameters(1).Value
        
        ' Insert Sub
        .CommandText = "xp_Outware_iOutwareSub"
        .Prepared = True

        For i = 0 To UBound(owSub)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, owSub(i).OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, ow.OutSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).OutSubSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).OrderSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).RollSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).BoxNo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, owSub(i).LotNo)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, owSub(i).OutQty)

            .Execute
            iSQL = iSQL + 1
        Next i
        
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COutWare.AddNewOutware"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(owSub) + 1)
    With ow
        sLog(0) = "xp_Outware_iOutware '" & .OrderID & "', " & .OutSeq & ", '" & .OutClss & "', '" & _
            .WorkID & "', " & .ExchRate & ", " & .UnitPrice & ", '" & .OutCustom & "', " & .LossRate & ", " & _
            .LossQty & ", " & .OutRoll & ", " & .OutQty & ", " & .OutRealQty & ", '" & .OutDate & "', '" & _
            .ResultDate & "', '" & .BoOutClss & "', '" & .OutTime & "', '" & .LoadTime & "', '" & .OutType & "', '" & _
            .Remark & "'"
    End With
    For i = 0 To UBound(owSub)
        With owSub(i)
            sLog(i + 1) = "xp_Outware_iOutwareSub '" & .OrderID & "', " & .OutSeq & ", " & .OutSubSeq & ", " & _
            .OrderSeq & ", " & .RollSeq & ", " & .BoxNo & ", '" & .LotNo & "', " & .OutQty & " "
        End With
    Next i
        
    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        AddNewOutware = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewOutware"), Err.Description
    Else
        Call LogService(sLog)

        AddNewOutware = True
    End If
End Function

Public Function UpdateOutware(ow As TOUTWARE, owSub() As TOUTWARESUB) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, j%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        ' Update
        .CommandText = "xp_Outware_uOutware"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, ow.OrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, ow.OutSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow.OutClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow.WorkID)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.ExchRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.UnitPrice)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, ow.OutCustom)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.LossRate)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, ow.LossQty)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, ow.OutRoll)
        .Parameters.Append .CreateParameter(, adDouble, adParamInput, 8, ow.OutQty)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 8, ow.OutRealQty)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, ow.OutDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, ow.ResultDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, ow.BoOutClss)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow.OutTime)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, ow.LoadTime)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, ow.Remark)

        .Execute
        iSQL = iSQL + 1

        ' Insert Sub
        .CommandText = "xp_Outware_iOutwareSub"
        .Prepared = True

        For i = 0 To UBound(owSub)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, owSub(i).OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, ow.OutSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).OutSubSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).OrderSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).RollSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, owSub(i).BoxNo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, owSub(i).LotNo)
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, owSub(i).OutQty)

            .Execute
            iSQL = iSQL + 1
        Next i
    
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COutWare.UpdateOutware"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(UBound(owSub) + 1)
    With ow
        sLog(0) = "xp_Outware_iOutware '" & .OrderID & "', " & .OutSeq & ", '" & .OutClss & "', '" & _
            .WorkID & "', " & .ExchRate & ", " & .UnitPrice & ", '" & .OutCustom & "', " & .LossRate & ", " & _
            .LossQty & ", " & .OutRoll & ", " & .OutQty & ", " & .OutRealQty & ", '" & .OutDate & "', '" & _
            .ResultDate & "', '" & .BoOutClss & "', '" & .OutTime & "', '" & .LoadTime & "', '" & .Remark & "'"
    End With
    For i = 0 To UBound(owSub)
        With owSub(i)
            sLog(i + 1) = "xp_Outware_iOutwareSub '" & .OrderID & "', " & .OutSeq & ", " & .OutSubSeq & ", " & _
            .OrderSeq & ", " & .RollSeq & ", " & .BoxNo & ", '" & .LotNo & "', " & .OutQty & " "
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateOutware = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateOutware"), Err.Description
    Else
        Call LogService(sLog)

        UpdateOutware = True
    End If
End Function

Public Function DeleteOutware(sOrderID As String, nOutSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    
    g_adoCon.BeginTrans
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_dOutware"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOutSeq)

        .Execute
        
''        .CommandText = "xp_Subul_dOutware"
''        .Prepared = True
''        Call ClearParameter(adoCmd)
''
''        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
''        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOutSeq)
''
''        .Execute
    End With
    
    g_adoCon.CommitTrans
    
    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COutWare.DeleteOutware"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(1)
    sLog(0) = "xp_Outware_dOutware '" & sOrderID & "', " & CStr(nOutSeq)
    sLog(1) = "xp_Subul_dOutware '" & sOrderID & "', " & CStr(nOutSeq)

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        DeleteOutware = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteOutware"), Err.Description
    Else
        Call LogService(sLog)

        DeleteOutware = True
    End If
End Function


Public Function UpdateBoOutClss(tOw() As TOUTWARE) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        .CommandText = "xp_Outware_uBoOutClss"
        .Prepared = True

        For i = 0 To UBound(tOw)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tOw(i).OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tOw(i).OutSeq)

            .Execute
            iSQL = iSQL + 1

            Call ClearParameter(adoCmd)
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COutWare.UpdateBoOutClss"
LogMessage:
    Set adoCmd = Nothing

    ' Logging
    ReDim sLog(UBound(tOw))
    For i = 0 To UBound(tOw)
        With tOw(i)
            sLog(i) = "xp_Outware_uBoOutClss '" & .OrderID & "', " & .OutSeq & " '"
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateBoOutClss = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateBoOutClss"), Err.Description
    Else
        Call LogService(sLog)

        UpdateBoOutClss = True
    End If
End Function

Public Function UpdateTranNo(sOrderID As String, nOutSeq As Integer, sTranNo As String, nTranSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Outware_uTranNO"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOutSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamOutput, 6, sTranNo)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, nTranSeq)
        
        .Execute
        
        sTranNo = IIf(IsNull(.Parameters(2).Value), "", .Parameters(2).Value)
        nTranSeq = IIf(IsNull(.Parameters(3).Value), 0, .Parameters(3).Value)
        
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "COutWare.UpdateTranNo"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    sLog(0) = "xp_Outware_uTranNO '" & sOrderID & "', " & CStr(nOutSeq)

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        UpdateTranNo = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateTranNO"), Err.Description
    Else
        Call LogService(sLog)

        UpdateTranNo = True
    End If
End Function

'--- 제품출고현황 ( 기간 )
Public Function GetOutWareView(ByVal sDate As String, ByVal eDate As String _
                            , ByVal nChkOrder As Integer, ByVal sOrderID As String _
                            , ByVal nChkCustom As Integer, ByVal sCustomID As String _
                            , ByVal nChkArticle As Integer, ByVal sArticleID As String _
                            , ByVal sTaxClss As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_OutWare_sOutWareView"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, eDate)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, sOrderID)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sTaxClss)

        Set GetOutWareView = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutWareView"), Err.Description
End Function


Public Function GetOutWareOrderByLot(ByVal sOrderID As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_OutWare_sOutWareOrderByLot"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)

        Set GetOutWareOrderByLot = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetOutWareOrderByLot"), Err.Description
End Function

