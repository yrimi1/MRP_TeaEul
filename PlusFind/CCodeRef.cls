VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCodeRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim SelData() As Variant

Public Property Get Data(Index As Integer)
    Data = SelData(Index)
End Property

Public Property Let Connection(adoConnect As Object)
    Set adoCon = adoConnect
End Property

'********************************************************
'* Date : 2000-07-24 (TUE)
'*
'* Description. Large   : 대분류
'*              Middle  : 중분류
'*              Item    : 찾을 코드, 명칭
'*
'********************************************************
Public Function Find(ByVal Large As Integer, Optional Middle, Optional Item) As Boolean
    Dim FindForm As Object
    Dim nPointer&, iLoop%, nStart%
    
    '------------------------------------------------------------------------'
    If Large = LG_ARTICLE Then
        Set FindForm = New frmArticle
        
        FindForm.optSize(0).Value = True
        Find = FindForm.SetMsg(SelData(), Item)
    '------------------------------------------------------------------------'
'    ElseIf Large = LG_WORK Then
'        Set FindForm = New frmOrderCode
'
'        Find = FindForm.SetMsg(SelData(), Item)
    '------------------------------------------------------------------------'
    ElseIf Large = LG_CUSTOM Then
        Set FindForm = New frmCustom
        
        FindForm.optSize(0).Value = True
        Find = FindForm.SetMsg(SelData(), Item)
    End If
    '========================================================================'
    
    Unload FindForm
    Set FindForm = Nothing
End Function

