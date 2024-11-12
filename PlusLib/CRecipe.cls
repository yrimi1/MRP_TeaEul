VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TRecipe
    OrderID    As String
    OrderSeq   As Integer
    RecipeSeq  As Integer
    ModifySeq  As Integer
    RecipeNO   As String
    RecipeDate As String
    PersonID   As String
    UnitWght   As Integer
    ChunkRate  As Single
    ModiClss   As String
    Qty    As Long
    Remark     As String
End Type

Public Type TRecipeSub
    OrderID    As String
    OrderSeq   As Integer
    RecipeSeq  As Integer
    ModifySeq  As Integer
    DyeAuxSeq  As Integer
    DyeAuxID   As String
    DyeAuxRate As Double
End Type


Public Type TMatch
    DyeID           As Long      ' 스케쥴 번호
    DyeSeq          As Integer      ' 탕수
    RecipeOrderID   As String   ' 처방전 관리번호
    RecipeOrderSeq  As Integer   ' 처방전 색상번호
    RecipeSeq       As Integer  ' 처방전 처방순위
    RecipeModifySeq As Integer  ' 처방전 변경순위
    PersonID        As String   ' 처방자
    Remark          As String   ' 비고 사항
End Type

Public Type TMatchSub
    DyeID       As Long     ' 스케쥴 번호
    DyeSeq      As Integer  ' 탕수
    DyeAuxSeq   As Integer  ' 염조제순위
    DyeAuxID    As String   ' 염조제 코드
    DyeAuxRate  As Double   ' 염조제 투입비율
    DyeAuxQty   As Single   ' 염조제 투입량
    DyeAuxRQty  As Single   ' 실제 투입량
End Type


Private Const CLASS_NAME As String = "CRecipe"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub



Public Function GetMaxRecipeNo() As String
    GetMaxRecipeNo = Format(GetMaxValue("Recipe", "RecipeNo"), "0000000000")
End Function




'**************************************************************************
'*
'* 처방전 내역 관리(2003-09-08)
'*
'* 최승백
'**************************************************************************

Public Function GetRecipe(nChkDate As Integer, sSDate As String, sEDate As String, nChkCustom As Integer, sCustom As String, _
    nChkOrder As Integer, sOrder As String, Optional nChkArticle As Integer = 0, Optional sArticle As String = " ") As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sRecipe"

        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sSDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustom)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkArticle)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticle)

        Set GetRecipe = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipe"), Err.Description
End Function



Public Function GetRecipeSub(sOrderID As String, nOrderSeq As Integer, nChkRecipeSeq As Integer, nRecipeSeq As Integer, _
        nChkKind As Integer, sKind As String, nModifySeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sRecipeSub"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkRecipeSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, CStr(nRecipeSeq))
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkKind)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sKind)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 3, CStr(nModifySeq))

        Set GetRecipeSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipeSubNEW"), Err.Description
End Function




Public Function GetRecipeHistory(sOrderID As String, nOrderSeq As Integer, nReworkSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sRecipeHistory"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nReworkSeq)

        Set GetRecipeHistory = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipeHistory"), Err.Description
End Function



Public Function DeleteRecipe(sOrderID As String, nOrderSeq As Integer, nRecipeSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_dRecipe"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nRecipeSeq)

        .Execute
    End With

    GoTo LogMessage

ErrHandler:
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CRecipe.DeleteRecipe"
LogMessage:
    Set adoCmd = Nothing

    ' Make Log Message
    ReDim sLog(0)
    sLog(0) = "xp_Recipe_dRecipe '" & sOrderID & "', '" & nOrderSeq & "', " & nRecipeSeq

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description)

        DeleteRecipe = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteRecipe"), Err.Description
    Else
        Call LogService(sLog)

        DeleteRecipe = True
    End If
End Function



Public Function GetRecipeOne(sRecipeNo As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sRecipeOne"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sRecipeNo)

        Set GetRecipeOne = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipeOneNEW"), Err.Description
End Function



Public Function AddNewRecipe(tRec As TRecipe, tRecSub() As TRecipeSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Insert Inspect
        .CommandText = "xp_Recipe_iRecipe"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tRec.OrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.OrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, tRec.RecipeSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.ModifySeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tRec.RecipeNO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tRec.RecipeDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tRec.PersonID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.UnitWght)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tRec.ChunkRate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tRec.ModiClss)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tRec.Qty)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 500, tRec.Remark)

        .Execute
        iSQL = iSQL + 1
        ' 재처방 순위 받음
        tRec.RecipeSeq = .Parameters(2).Value

        ' Insert InspectSub
        .CommandText = "xp_Recipe_iRecipeSub"
        .Prepared = True

        For i = 0 To UBound(tRecSub)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tRec.OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.OrderSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.RecipeSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.ModifySeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRecSub(i).DyeAuxSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tRecSub(i).DyeAuxID)
            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 5, tRecSub(i).DyeAuxRate)

            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CRecipe.AddNewRecipe"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(tRecSub) + 1)
    With tRec
        sLog(0) = "xp_Recipe_iRecipeNEW '" & .OrderID & "', '" & .OrderSeq & "', " & .RecipeSeq & ", " & .ModifySeq & ", '" & _
            .RecipeNO & "', '" & .RecipeDate & "', '" & .PersonID & "'"
    End With
    For i = 0 To UBound(tRecSub)
        With tRecSub(i)
            sLog(i + 1) = "xp_Recipe_iRecipeSubNEW '" & tRec.OrderID & "', '" & tRec.OrderSeq & "', " & tRec.RecipeSeq & ", " & _
                tRec.ModifySeq & ", " & .DyeAuxSeq & ", '" & .DyeAuxID & "', " & .DyeAuxRate
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        AddNewRecipe = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewRecipe"), Err.Description
    Else
        Call LogService(sLog)

        AddNewRecipe = True
    End If
End Function



Public Function UpdateRecipe(tRec As TRecipe, tRecSub() As TRecipeSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Update Inspect
        .CommandText = "xp_Recipe_uRecipe"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tRec.OrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.OrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.RecipeSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, tRec.ModifySeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tRec.RecipeNO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tRec.RecipeDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tRec.PersonID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tRec.UnitWght)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tRec.ChunkRate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, tRec.ModiClss)
        .Parameters.Append .CreateParameter(, adSingle, adParamInput, 4, tRec.Qty)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 500, tRec.Remark)

        .Execute
        ' 처방전 변경순위 받음
        tRec.ModifySeq = .Parameters(3).Value
        
        iSQL = iSQL + 1

        ' Insert InspectSub
        .CommandText = "xp_Recipe_iRecipeSub"
        .Prepared = True

        For i = 0 To UBound(tRecSub)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tRec.OrderID)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.OrderSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.RecipeSeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRec.ModifySeq)
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tRecSub(i).DyeAuxSeq)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tRecSub(i).DyeAuxID)
            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 5, tRecSub(i).DyeAuxRate)

            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CRecipe.UpdateRecipe"
LogMessage:
    Set adoCmd = Nothing

    ReDim sLog(UBound(tRecSub) + 1)
    With tRec
        sLog(0) = "xp_Recipe_uRecipeNEW '" & .OrderID & "', '" & .OrderSeq & "', " & .RecipeSeq & ", " & .ModifySeq & ", '" & _
            .RecipeNO & "', '" & .RecipeDate & "', '" & .PersonID & "', " & .UnitWght & ", '" & .Remark & "'"
    End With
    For i = 0 To UBound(tRecSub)
        With tRecSub(i)
            sLog(i + 1) = "xp_Recipe_iRecipeSubNEW '" & tRec.OrderID & "', '" & tRec.OrderSeq & "', " & tRec.RecipeSeq & ", " & _
                tRec.ModifySeq & ", " & .DyeAuxSeq & ", '" & .DyeAuxID & "', " & .DyeAuxRate
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateRecipe = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "UpdateRecipe"), Err.Description
    Else
        Call LogService(sLog)

        UpdateRecipe = True
    End If
End Function



Public Function GetRecipeCount(sOrderID As String, nOrderSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sRecipeCount"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        
        Set GetRecipeCount = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipeCount"), Err.Description
End Function



Public Function GetRecipeSubAll(sOrderID As String, nOrderSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sRecipeSubAll"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        
        Set GetRecipeSubAll = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipeSubAll"), Err.Description
End Function


Public Function GetRecipeSubByRecipeSeq(sOrderID As String, nOrderSeq As Integer, nChkRecipeSeq As Integer, nRecipeSeq As Integer, _
        nChkKind As Integer, sKind As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Dye_sRecipeSub"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkRecipeSeq)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, CStr(nRecipeSeq))
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkKind)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, sKind)
        
        Set GetRecipeSubByRecipeSeq = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipeSubByRecipeSeq"), Err.Description
End Function


Public Function GetRecipeByColor(sOrderID As String, nOrderSeq As Integer, nChkReworkSeq As Integer, nReworkSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sRecipeByColor"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkReworkSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nReworkSeq)
        
        
        Set GetRecipeByColor = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipeByColor"), Err.Description
End Function



Public Function IsExistRecipe(sOrderID As String, nOrderSeq As Integer) As Boolean
    IsExistRecipe = IsExistRecord("Recipe", "OrderID = '" & sOrderID & "' AND OrderSeq = '" & nOrderSeq & "'")
    
End Function



Public Function GetRecipeList(sDate As String, sEDate As String, nChkCustom As Integer, sCustomID As String, nChkOrder As Integer, sOrder As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_pRecipeListNEW"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sDate)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sEDate)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkCustom)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sOrder)
        
        Set GetRecipeList = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipeList"), Err.Description
End Function



Public Function GetRecipeUseCount(sOrderID As String, nOrderSeq As Integer, nRecipeSeq As Integer, nCount As Integer) As Integer
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sRecipeUseInMatch"

        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, sOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nOrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nRecipeSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamOutput, 2, nCount)
        
        .Execute
        
        GetRecipeUseCount = .Parameters(3).Value
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing
    GetRecipeUseCount = 0
    
    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRecipeUseCount"), Err.Description
End Function

' *****************************************************************
' *
' *  염색 지시내역 불러오기
' *
' * frogfinger - frmRecipeCalc
' *****************************************************************

Public Function GetDyeCommandOne(nID As Long, nSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Dye_sCommandOne"

        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nSeq)
        
        Set GetDyeCommandOne = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetDyeCommandOne"), Err.Description
End Function



' *****************************************************************
' *
' *  염색 지시별 카드 상세내역
' *
' * frogfinger - frmRecipeCalc
' *****************************************************************

Public Function GetRapidCommandSub(nDyeID As Long, nDyeSeq As Integer)
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Dye_sCommandDyeSub"

        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nDyeID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nDyeSeq)
        
        Set GetRapidCommandSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetRapidCommandSub"), Err.Description

End Function


'**************************************************************************
'*
'* 평량지시 신규 작성 (2003-12-02)
'*
'* Author : 최승백
'**************************************************************************
Public Function AddNewMatchData(tData As TMatch, tDataSub() As TMatchSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Insert Inspect
        .CommandText = "xp_Dye_iMatch"

        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, tData.DyeID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.DyeSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tData.RecipeOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.RecipeOrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.RecipeSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.RecipeModifySeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tData.PersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 500, tData.Remark)
                
        .Execute
        iSQL = iSQL + 1

        .CommandText = "xp_Dye_iMatchSub"
        .Prepared = True

        For i = 0 To UBound(tDataSub)
            Call ClearParameter(adoCmd)
    
            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, tData.DyeID)      ' 스케쥴 번호
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.DyeSeq)    ' 차수
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tDataSub(i).DyeAuxSeq)   ' 염료숮위
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tDataSub(i).DyeAuxID)    ' 염료ID
            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 5, tDataSub(i).DyeAuxRate)    ' 투입비율
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tDataSub(i).DyeAuxQty)     ' 처방량
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tDataSub(i).DyeAuxRQty)
                    
            .Execute
            iSQL = iSQL + 1
        Next i
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CRecipe.AddNewMatchData"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(tDataSub) + 1)
    With tData
        sLog(0) = "xp_Dye_iMatch '" & .DyeID & "', " & .DyeSeq & ", " & _
            .RecipeOrderID & "', '" & .RecipeOrderSeq & ", " & .RecipeSeq & ", " & .RecipeModifySeq & ", '" & _
            .PersonID
            
    End With
    For i = 0 To UBound(tDataSub)
        With tDataSub(i)
            sLog(i + 1) = "xp_Dye_iMatchSub '" & .DyeID & "', " & .DyeSeq & ", " & .DyeAuxSeq & ", " & _
                .DyeAuxID & ", '" & .DyeAuxRate & ", " & .DyeAuxQty & ", " & .DyeAuxRQty
                
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        AddNewMatchData = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "AddNewMatchData"), Err.Description
    Else
        Call LogService(sLog)

        AddNewMatchData = True
    End If
End Function




'**************************************************************************
'*
'* 평량지시 변경 (2003-12-02)
'*
'* Author : 최승백
'**************************************************************************
Public Function UpdateMatchData(tData As TMatch, tDataSub() As TMatchSub) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Insert Inspect
        .CommandText = "xp_Dye_uModifyMatch"

        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, tData.DyeID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.DyeSeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 10, tData.RecipeOrderID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.RecipeOrderSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.RecipeSeq)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.RecipeModifySeq)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, tData.PersonID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 500, tData.Remark)
                
        .Execute
        iSQL = iSQL + 1

        .CommandText = "xp_Dye_iMatchSub"
        .Prepared = True

        For i = 0 To UBound(tDataSub)
            Call ClearParameter(adoCmd)

            .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, tData.DyeID)      ' 스케쥴 번호
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tData.DyeSeq)    ' 차수
            .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, tDataSub(i).DyeAuxSeq)   ' 염료숮위
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, tDataSub(i).DyeAuxID)    ' 염료ID
            .Parameters.Append .CreateParameter(, adDouble, adParamInput, 5, tDataSub(i).DyeAuxRate)    ' 투입비율
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tDataSub(i).DyeAuxQty)     ' 처방량
            .Parameters.Append .CreateParameter(, adSingle, adParamInput, 5, tDataSub(i).DyeAuxRQty)
                    
            .Execute
            iSQL = iSQL + 1
        Next i
        
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CRecipe.ModifyMatchData"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(UBound(tDataSub) + 1)
    With tData
        sLog(0) = "xp_Rapid_uModifyMatch '" & .DyeID & "', " & .DyeSeq & ", " & _
            .RecipeOrderID & "', '" & .RecipeOrderSeq & ", " & .RecipeSeq & ", " & .RecipeModifySeq & ", '" & _
            .PersonID
            
    End With
    For i = 0 To UBound(tDataSub)
        With tDataSub(i)
            sLog(i + 1) = "xp_Dye_iMatchSub '" & .DyeID & "', " & .DyeSeq & ", " & .DyeAuxSeq & ", " & _
                .DyeAuxID & ", '" & .DyeAuxRate & ", " & .DyeAuxQty & ", " & .DyeAuxRQty
                
        End With
    Next i

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        UpdateMatchData = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "ModifyMatchData"), Err.Description
    Else
        Call LogService(sLog)

        UpdateMatchData = True
    End If
End Function




'**************************************************************************
'*
'* 평량지시 내역 확인 (2003-12-02)
'*
'* Author : 최승백
'**************************************************************************
Public Function GetMatch(nDyeID As Long, nDyeSeq As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Dye_sMatch"

        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nDyeID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nDyeSeq)
        
        Set GetMatch = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetMatch"), Err.Description
End Function



'**************************************************************************
'*
'* 평량지시 세부내역 확인 (2003-12-02)
'*
'* Author : 최승백
'**************************************************************************
Public Function GetMatchSub(nDyeID As Long, nDyeSeq As Integer, sKind As String) As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Dye_sMatchSub"

        .Parameters.Append .CreateParameter(.CommandText, adInteger, adParamInput, 4, nDyeID)
        .Parameters.Append .CreateParameter(.CommandText, adSmallInt, adParamInput, 2, nDyeSeq)
        .Parameters.Append .CreateParameter(.CommandText, adVarChar, adParamInput, 1, sKind)
        
        Set GetMatchSub = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "GetMatchSub"), Err.Description
End Function



'**************************************************************************
'*
'* 평량지시 삭제 (2003-12-02)
'*
'* Author : 최승백
'**************************************************************************
Public Function DeleteMatchData(nDyeID As Long, nDyeSeq As Integer) As Boolean
    Dim adoCmd As ADODB.Command
    Dim sLog() As String, bError As Boolean
    Dim i%, iSQL%

    On Error GoTo ErrHandler

    g_adoCon.BeginTrans

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc

        ' Delete Match
        .CommandText = "xp_Dye_dMatch"

        .Parameters.Append .CreateParameter(, adInteger, adParamInput, 4, nDyeID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nDyeSeq)
                
        .Execute
        iSQL = iSQL + 1
        
    End With

    g_adoCon.CommitTrans

    GoTo LogMessage

ErrHandler:
    g_adoCon.RollbackTrans
    bError = True
    MsgBox Err.Number & " / " & Err.Description, vbCritical, "CRecipe.DeleteMatchData"
LogMessage:
    Set adoCmd = Nothing

    ' Logging Error
    ReDim sLog(0)
    
    sLog(0) = "xp_Dye_dMatch '" & nDyeID & "', " & nDyeSeq
            

    If bError Then
        Call ErrLogService(sLog, Err.Number, Err.Description, iSQL)

        DeleteMatchData = False
''        Err.Raise Err.Number, ErrorSource(CLASS_NAME, "DeleteMatchData"), Err.Description
    Else
        Call LogService(sLog)

        DeleteMatchData = True
    End If
End Function

Public Function GetCardList(nChkCustomID As Integer, sCustomID As String, _
                    nChkArticleID As Integer, sArticleID As String, _
                    nChkOrder As Integer, sOrder As String, _
                    nChkCardID As Integer, sCardID As String, sSplitID, _
                    nChkProcessID As Integer, sProcessID As String, _
                    nChkRecipeFlag As Integer) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sCard"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCardID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sSplitID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkProcessID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, sProcessID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 1, nChkRecipeFlag)

        Set GetCardList = .Execute
    End With
    
    Set adoCmd = Nothing
End Function

Public Function GetRecipeCalcList(nChkCustomID As Integer, sCustomID As String, _
                    nChkArticleID As Integer, sArticleID As String, _
                    nChkOrder As Integer, sOrder As String, _
                    nChkCardID As Integer, sCardID As String, sSplitID) As Recordset
    Dim adoCmd As ADODB.Command

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Recipe_sRecipeCalcList"
        
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCustomID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sCustomID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkArticleID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 4, sArticleID)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkOrder)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 24, sOrder)
        .Parameters.Append .CreateParameter(, adSmallInt, adParamInput, 2, nChkCardID)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 8, sCardID)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 4, sSplitID)

        Set GetRecipeCalcList = .Execute
    End With
    
    Set adoCmd = Nothing
End Function


