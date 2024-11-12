VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "CLogin"

Public Property Let Connection(adoConnect As Object)
    Set g_adoCon = adoConnect
End Property

Public Property Let UserName(sUserName As String)
    g_sUserName = sUserName
End Property

Private Sub Class_Terminate()
    If Not g_adoCon Is Nothing Then Set g_adoCon = Nothing
End Sub

Public Function Login(sUserID As String, Optional sPassword As String) As String
    Dim adoCmd As ADODB.Command
    Dim rs     As ADODB.Recordset
    Dim sSQL$

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Common_Login"

        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 15, sUserID)

        Set rs = .Execute
    End With
    Set adoCmd = Nothing
        
    If rs.EOF Then
        Err.Raise vbObjectError + 1010, ErrorSource(CLASS_NAME, "Login"), LoadResString(1010)
    Else
        If sPassword <> rs!Password Then
            Err.Raise vbObjectError + 1020, ErrorSource(CLASS_NAME, "Login"), LoadResString(1020)
        End If
    End If
    Login = rs!Name
    rs.Close
    Set rs = Nothing

    Exit Function

ErrHandler:
    Set rs = Nothing
    Set adoCmd = Nothing

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "Login"), Err.Description
End Function

Public Function GetNow() As Variant
    Dim adoCmd As ADODB.Command

    On Error GoTo ErrHandler

    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = g_adoCon
        .CommandType = adCmdStoredProc
        .CommandText = "xp_Common_sNow"

        .Parameters.Append .CreateParameter(, adDate, adParamOutput, 8, 0)
        .Execute

        GetNow = .Parameters(0).Value
    End With
    Set adoCmd = Nothing

    Exit Function

ErrHandler:
    Set adoCmd = Nothing
    GetNow = Null

    Err.Raise Err.Number, ErrorSource(CLASS_NAME, "Login"), Err.Description
End Function
