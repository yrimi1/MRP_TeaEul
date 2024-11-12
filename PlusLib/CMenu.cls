VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TUSERMENU
    sPersonID   As String
    sMenuID     As String
    nSeq        As Integer
    nLevel      As Integer
    sParentID   As String
    sPrintClss  As String
    sSelectClss As String
    sAddNewClss As String
    sUpdateClss As String
    sDeleteClss As String
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

Public Function GetMainMenu() As Recordset
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Code_sMenu"

        Set GetMainMenu = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set GetMainMenu = Nothing
    Set adoCmd = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetUserMenu(sUserID As String) As Recordset
    Dim adoCmd As ADODB.Command
    Dim SQL$

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Menu_sUserMenu"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, sUserID)

        Set GetUserMenu = .Execute
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set GetUserMenu = Nothing
    Set adoCmd = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function DeleteUserMenu(sPersonID As String) As Boolean
    Dim SQL As String

    SQL = "DELETE FROM [MenuUser] WHERE PersonID = '" & sPersonID & "'"

    DeleteUserMenu = HandleDB(SQL)
End Function

