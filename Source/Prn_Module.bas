Attribute VB_Name = "Prn_Module"


Option Explicit

'--- �ŷ����� Detail ��¿� ����ü
Type TStuffDetail
    Color       As String    'Color
    LotNo       As String    'LotNo
    nColorRoll  As Integer   'Color�� �� �հ�
    nColorQty   As Integer   'Color�� �հ�
    sRollStr    As String    '�����
    nTotQty     As Integer   '��
End Type

Type TRoll
    RollNo      As Integer   ' Roll Number
    ColorID     As String    ' �����ڵ�
    Color       As String    ' �����
    Qty         As Integer   ' �� ����
End Type

Dim dXOffSet As Long, dYOffSet As Long, dyPos As Long

'''Public Sub SetStuffINReturnGoods(ByVal sStuffDate As String, ByVal sStuffClss As String, ByVal nStuffSeq As Integer, ByVal oFlex As VSFlexGrid)
'''    Dim oStuffIn As Pluslib2.cStuffIN
'''    Dim RsHeader As ADODB.Recordset, RsDetail As ADODB.Recordset
'''    Dim rsData As ADODB.Recordset
'''    Dim nRollvar(), nCols As Integer
'''
'''    On Error GoTo ErrHandler
'''
'''    Set oStuffIn = New Pluslib2.cStuffIN
'''    oStuffIn.Connection = g_adoCon
'''    oStuffIn.UserName = g_sUserName
'''
'''    If oStuffIn.GetStuffINReturnGoods(sStuffDate, sStuffClss, nStuffSeq, RsHeader) Then
'''        Call SetPrint(RsHeader, oFlex)
'''        Set oStuffIn = Nothing
'''    End If
'''
'''    Exit Sub
'''ErrHandler:
'''    MsgBox ("��ǰ ���� ��� �� ���� �߻� ")
'''End Sub




'Public Sub PrnData(ByVal xPos As Long, ByVal yPos As Long, ByVal dStr As String)
'    Printer.CurrentX = xPos
'    Printer.CurrentY = yPos
'    Printer.Print Trim$(dStr)
'End Sub
'
'Public Function PrintDot(nXPos As Integer, nYPos As Integer, sStr As String, Optional nFont As Integer = 10)
'    With Printer
'        .CurrentX = nXPos
'        .CurrentY = nYPos
'        .Font.Size = nFont
'    End With
'    Printer.Print sStr
'End Function

''Public Function ChkNullValue(ChkCol As Field) As Variant
''
''
''    If IsNull(ChkCol) Then
''        If ChkCol.Type = adChar Or ChkCol.Type = adVarChar Then
''            ChkNullValue = ""
''        Else
''            ChkNullValue = 0
''        End If
''    Else
''        If ChkCol.Type = adChar Or ChkCol.Type = adVarChar Then
''            ChkNullValue = Trim$(ChkCol)
''        Else
''            ChkNullValue = ChkCol
''        End If
''    End If
''End Function





