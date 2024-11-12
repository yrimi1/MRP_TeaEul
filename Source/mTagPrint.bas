Attribute VB_Name = "TagPrint"
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Const FONT_NAME  As String = "°ß°íµñ"

Private Enum EnumItem
    IO_DATA = 0
    IO_BARCODE = 1
    IO_TEXT = 2
    IO_LINE = 3
    IO_RECT = 4
    IO_DIAMOND = 5
    IO_CIRCLE = 6
    IO_IMAGE = 7
End Enum

Private Type TTag
    sTagID        As String
    sTag          As String
    nWidth        As Long
    nHeight       As Long
    nClss         As Integer
    nDefectClss   As Integer
    nDefHeight    As Integer
    nDefBaseY     As Integer
    nDefBaseX1    As Integer
    nDefBaseX2    As Integer
    nDefBaseX3    As Integer
    nDefGapY      As Integer
    nDefGapX1     As Integer
    nDefGapX2     As Integer
    nDefLength    As Integer
    nDefHCount    As Integer
    nDefBarClss   As Integer
    nGap          As Integer
    sDirect       As String
End Type

Private Type TTagSub
    sName         As String
    nType         As Integer
    nAlign        As Integer
    x             As Integer
    y             As Integer
    nFont         As Integer
    nLength       As Integer
    nHMulti       As Integer
    nVMulti       As Integer
    nRotation     As Integer
    sText         As String
    nSpace        As Integer
    nRelation     As Integer
    nPrevItem     As Integer
    nBarType      As Integer
    nBarHeight    As Integer
    nFigureWidth  As Integer
    nFigureHeight As Integer
    nThickness    As Integer
    sImageFile    As String
    nWidth        As Integer
    nHeight       As Integer
    nVisible      As Integer
End Type

Private Const DPI_RATIO As Single = 0.8
Private Const GAP_X     As Integer = 4
Private Const GAP_Y     As Integer = 2

Private m_tTag    As TTag
Private m_tItem() As TTagSub

Private m_sData()   As String
Private m_picPrint  As PictureBox
Private m_picFont   As PictureBox

' vDefect(i, 0) = TagName
' vDefect(i, 1) = Position
' vDefect(i, 2) = Demerit

Public Function MakeCleverTagPrintString(vData As Variant, ByVal sTagID As String, Optional vDefect As Variant, Optional nTotalLength As Integer, Optional nPrintCount As Integer = 1, Optional nDefectCnt As Integer = 1) As String
    Dim oCode As PlusLib2.CCode
    Dim oTag  As PlusLib2.CTag
    Dim rs     As ADODB.Recordset
    Dim rsTag  As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    Dim i&, j&, k&, l&, nLen&, x&, y&, nPixWidth%, nPixHeight%, nPixel&
    Dim sText$, sChar$, nBitmap As Byte
    Dim sBarType(1) As String

'    On Error GoTo ErrHandler

    ' Data Loads ***********************************************************************************
    Set oTag = New PlusLib2.CTag
    oTag.Connection = g_adoCon
    
    Set rsTag = oTag.GetTag(1, sTagID)
    Set rsItem = oTag.GetTagSub(sTagID)

    With m_tTag
        .sTag = rsTag!Tag
        .nWidth = rsTag!Width
        .nHeight = rsTag!Height
        .nClss = rsTag!Clss
        .nDefectClss = rsTag!DefectClss
        .nDefHeight = rsTag!DefHeight
        .nDefBaseY = rsTag!DefBaseY
        .nDefBaseX1 = rsTag!DefBaseX1
        .nDefBaseX2 = rsTag!DefBaseX2
        .nDefBaseX3 = rsTag!DefBaseX3
        .nDefGapY = rsTag!DefGapY
        .nDefGapX1 = rsTag!DefGapX1
        .nDefGapX2 = rsTag!DefGapX2
        .nDefLength = rsTag!DefLength
        .nDefHCount = rsTag!DefHCount
        .nDefBarClss = rsTag!DefBarClss
        .nGap = rsTag!Gap
        .sDirect = rsTag!Direct
    End With
    rsTag.Close
    Set rsTag = Nothing

    Do Until rsItem.EOF
        ReDim Preserve m_tItem(i)

        With m_tItem(i)
            .sName = rsItem!Name
            .nType = rsItem!Type
            .nAlign = rsItem!Align
            .x = rsItem!x
            .y = rsItem!y
            .nFont = rsItem!Font
            .nLength = rsItem!Length
            .nHMulti = rsItem!HMulti
            .nVMulti = rsItem!VMulti
            .nRotation = rsItem!Rotation
            .sText = rsItem!Text
            .nSpace = rsItem!Space
            .nRelation = rsItem!Relation
            .nPrevItem = rsItem!PrevItem
            .nBarType = rsItem!BarType
            .nBarHeight = rsItem!BarHeight
            .nFigureWidth = rsItem!FigureWidth
            .nFigureHeight = rsItem!FigureHeight
            .nThickness = rsItem!Thickness
            .sImageFile = rsItem!ImageFile
            .nWidth = rsItem!Width
            .nHeight = rsItem!Height
            .nVisible = rsItem!Visible
        End With

        i = i + 1
        rsItem.MoveNext
    Loop
    rsItem.Close
    Set rsItem = Nothing

    ' Make Print Head ******************************************************************************
    With m_tTag
        MakeCleverTagPrintString = "SPEED 3" & vbCr & vbLf & _
            "DENSITY 11" & vbCr & vbLf & _
            "SET CUTTER OFF" & vbCr & vbLf & _
            "SET PEEL OFF" & vbCr & vbLf & _
            "DIRECTION " & .sDirect & vbCr & vbLf & _
            "SIZE " & SetCurrency(.nWidth / 10, 1) & " mm, " & SetCurrency(.nHeight / 10, 1) & " mm" & vbCr & vbLf & _
            "GAP " & SetCurrency(.nGap / 10, 1) & " mm, 0 mm" & vbCr & vbLf & _
            "OFFSET 0.0 mm" & vbCr & vbLf & _
            "REFERENCE 0, 0" & vbCr & vbLf & _
            "CLS" & vbCr & vbLf
    End With

    Set m_picPrint = frmInspect.picPrint

    ' Make Print Fields ****************************************************************************
    ReDim m_sData(UBound(vData))
    For i = 0 To UBound(vData)
        m_sData(i) = vData(i)
    Next i

    For i = 0 To UBound(m_tItem)
        With m_tItem(i)
            If .nVisible > 0 Then
                If .nType = IO_BARCODE Then
                    If .nPrevItem = 0 Then
                        If .nBarType = 0 Then   ' 1:1 Code
                            sBarType(0) = "1"
                            sBarType(1) = "1"
                        Else                    ' 2:5 Code
                            sBarType(0) = "2"
                            sBarType(1) = "5"
                        End If
    
                        If .nRotation = 0 Then  ' 0 Angle
                            MakeCleverTagPrintString = MakeCleverTagPrintString & "BARCODE " & _
                                CStr(GAP_X + CInt(.x * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & ", " & _
                                Chr(34) & "128" & Chr(34) & ", " & CStr(CInt(.nBarHeight * DPI_RATIO)) & ", 0, 0, " & _
                                sBarType(0) & ", " & sBarType(1) & ", " & Chr(34) & GetBarCodeItemText(i) & Chr(34) & vbCr & vbLf
                        Else                    ' 90 Angle
                            MakeCleverTagPrintString = MakeCleverTagPrintString & "BARCODE " & _
                                CStr(GAP_X + CInt((.x + .nBarHeight) * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & ", " & _
                                Chr(34) & "128" & Chr(34) & ", " & CStr(CInt(.nBarHeight * DPI_RATIO)) & ", 0, 90, " & _
                                sBarType(0) & ", " & sBarType(1) & ", " & Chr(34) & GetBarCodeItemText(i) & Chr(34) & vbCr & vbLf
                        End If
                    End If
                ElseIf .nType = IO_DATA Or .nType = IO_TEXT Then
                    Dim nFontDot%, nGapY%, nHMulti%, nVMulti%

                    sText = GetItemText(i)
                    nHMulti = .nHMulti + 1
                    nVMulti = .nVMulti + 1
                    nFontDot = GetCleverFontDot(.nFont)
                    nGapY = GetCleverFontGapY(.nFont) / nVMulti

                    If IsHangul(sText) Then
                        If .nRotation = 0 Then  ' 0 Angle
                            x = GAP_X + CInt(.x * DPI_RATIO)
                            nLen = Len(sText) - 1
                            For j = 0 To nLen
                                sChar = Mid(sText, j + 1, 1)
                                If IsHangul(sChar) Then
                                    MakeCleverTagPrintString = MakeCleverTagPrintString & "BITMAP " & CStr(x) & ", " & CStr(GAP_Y + CInt(.y * DPI_RATIO) + nGapY) & ", " & _
                                        ArithUpper(CSng((nFontDot * nHMulti * DPI_RATIO) / 8)) & ", " & ArithUpper(CSng(nFontDot * nVMulti * DPI_RATIO)) & ", 1," & _
                                        GetHangulBitmap(0, .nFont, sChar, .nRotation, nHMulti, nVMulti) & vbCr & vbLf
                                    x = x + (nFontDot * nHMulti - .nFont * 2)
                                Else
                                    MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & CStr(x) & ", " & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & ", " & _
                                        Chr(34) & CStr(.nFont) & Chr(34) & ", " & "0, " & CStr(nHMulti) & ", " & CStr(nVMulti + 1) & ", " & _
                                        Chr(34) & sChar & Chr(34) & vbCr & vbLf
                                    x = x + (nFontDot * nHMulti / 3)
                                End If
                            Next j
                        Else                    ' 90 Angle
                            y = GAP_Y + CInt(.y * DPI_RATIO)
                            nLen = Len(sText) - 1
                            For j = 0 To nLen
                                sChar = Mid(sText, j + 1, 1)
                                If IsHangul(sChar) Then
                                    MakeCleverTagPrintString = MakeCleverTagPrintString & "BITMAP " & CStr(GAP_X + CInt(.x * DPI_RATIO) - (nFontDot / 5)) & ", " & CStr(y) & ", " & _
                                        ArithUpper(CSng((nFontDot * nVMulti * DPI_RATIO) / 8)) & ", " & ArithUpper(CSng(nFontDot * nHMulti * DPI_RATIO)) & ", 1," & _
                                        GetHangulBitmap(0, .nFont, sChar, .nRotation, nHMulti, nVMulti) & vbCr & vbLf
                                    y = y + (nFontDot * nHMulti - .nFont * 2)
                                Else
                                    MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & CStr(GAP_X + CInt(.x * DPI_RATIO) + (nFontDot / 2) * (nVMulti + 1)) & ", " & CStr(y) & ", " & _
                                        Chr(34) & CStr(.nFont) & Chr(34) & ", " & "90, " & CStr(nVMulti + 1) & ", " & CStr(nHMulti) & ", " & _
                                        Chr(34) & sChar & Chr(34) & vbCr & vbLf
                                    y = y + (nFontDot * nHMulti / 3)
                                End If
                            Next j
                        End If
                    Else
                        If .nRotation = 0 Then  ' 0 Angle
                            MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & _
                                CStr(GAP_X + CInt(.x * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & ", " & _
                                Chr(34) & CStr(.nFont) & Chr(34) & ", " & "0, " & CStr(nHMulti) & ", " & CStr(nVMulti + 1) & ", " & _
                                Chr(34) & sText & Chr(34) & vbCr & vbLf
                        Else                    ' 90 Angle
                            MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & _
                                CStr(GAP_X + CInt(.x * DPI_RATIO) + (nFontDot / 2) * (nVMulti + 1)) & ", " & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & ", " & _
                                Chr(34) & CStr(.nFont) & Chr(34) & ", " & "90, " & CStr(nVMulti + 1) & ", " & CStr(nHMulti) & ", " & _
                                Chr(34) & sText & Chr(34) & vbCr & vbLf
                        End If
                    End If
                ElseIf .nType = IO_LINE And (.nFigureHeight <= 5 Or .nFigureWidth <= 5) Then
                    j = GAP_X + CInt(.x * DPI_RATIO)
                    k = GAP_Y + CInt(.y * DPI_RATIO)

                    MakeCleverTagPrintString = MakeCleverTagPrintString & "BOX " & CStr(j) & ", " & CStr(k) & ", " & CStr(j + CInt(.nFigureWidth * DPI_RATIO)) & ", " & CStr(k + CInt(.nFigureHeight * DPI_RATIO)) & ", " & CStr(.nThickness) & vbCr & vbLf
                ElseIf .nType = IO_LINE Or .nType = IO_DIAMOND Or .nType = IO_IMAGE Then
                    Dim nStart%, nVWidth%, nVHeight%, nAdjust#

                    m_picPrint.Width = ArithUpper(.nFigureWidth * DPI_RATIO) * Screen.TwipsPerPixelX + (Screen.TwipsPerPixelX * 2)
                    m_picPrint.Height = ArithUpper(.nFigureHeight * DPI_RATIO) * Screen.TwipsPerPixelY + (Screen.TwipsPerPixelY * 2)
                    m_picPrint.ScaleWidth = m_picPrint.Width - Screen.TwipsPerPixelX * 2
                    m_picPrint.ScaleHeight = m_picPrint.Height - Screen.TwipsPerPixelY * 2

                    m_picPrint.Picture = Nothing
                    m_picPrint.Cls

                    nAdjust = .nThickness / 2
                    nStart = Screen.TwipsPerPixelX * ArithLower(nAdjust)
                    nVWidth = m_picPrint.ScaleWidth - Screen.TwipsPerPixelX * ArithUpper(nAdjust)
                    nVHeight = m_picPrint.ScaleHeight - Screen.TwipsPerPixelY * ArithUpper(nAdjust)

                    m_picPrint.DrawWidth = .nThickness
                    If .nType = IO_LINE Then
                        If .nAlign = 0 Then ' like \
                            m_picPrint.Line (0, 0)-(m_picPrint.ScaleWidth, m_picPrint.ScaleHeight)
                        Else                ' like /
                            m_picPrint.Line (m_picPrint.ScaleWidth, 0)-(0, m_picPrint.ScaleHeight)
                        End If
                    ElseIf .nType = IO_RECT Then
                        m_picPrint.Line (nStart, nStart)-(nVWidth, nStart)
                        m_picPrint.Line -(nVWidth, nVHeight)
                        m_picPrint.Line -(nStart, nVHeight)
                        m_picPrint.Line -(nStart, nStart)
                    ElseIf .nType = IO_DIAMOND Then
                        m_picPrint.Line (nVWidth / 2, nStart)-(nVWidth, nVHeight / 2)
                        m_picPrint.Line -(nVWidth / 2, nVHeight)
                        m_picPrint.Line -(nStart, nVHeight / 2)
                        m_picPrint.Line -(nVWidth / 2, nStart)
                    ElseIf .nType = IO_CIRCLE Then
                        m_picPrint.Circle (ArithLower(m_picPrint.ScaleWidth / 2), ArithLower(m_picPrint.ScaleHeight / 2)), _
                            ArithLower(IIf(m_picPrint.ScaleWidth > m_picPrint.ScaleHeight, nVWidth, nVHeight) / 2 - Screen.TwipsPerPixelX), , , , _
                            CDbl(m_picPrint.ScaleHeight / m_picPrint.ScaleWidth)
                    ElseIf .nType = IO_IMAGE Then
                        m_picPrint.Picture = LoadPicture(.sImageFile)
                    End If

                    MakeCleverTagPrintString = MakeCleverTagPrintString & "BITMAP " & CStr(GAP_X + CInt(.x * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & ", " & _
                        ArithUpper(CSng((.nFigureWidth * DPI_RATIO) / 8)) & ", " & ArithUpper(CSng(.nFigureHeight * DPI_RATIO)) & ", 1,"

                    nPixWidth = m_picPrint.ScaleWidth / Screen.TwipsPerPixelX - 1
                    nPixHeight = m_picPrint.ScaleHeight / Screen.TwipsPerPixelY - 1
                    For j = 0 To nPixHeight
                        nBitmap = 255
                        For k = 0 To nPixWidth
                            l = k Mod 8

                            nPixel = GetPixel(m_picPrint.hdc, k, j)
                            If nPixel = 0 Then
                                nBitmap = nBitmap - (2 ^ Abs(l - 7))
                            End If

                            If l = 7 Or k = nPixWidth Then
                                MakeCleverTagPrintString = MakeCleverTagPrintString & ChrB(nBitmap) & ChrB(0)
                                nBitmap = 255
                            End If
                        Next k
                    Next j

                    MakeCleverTagPrintString = MakeCleverTagPrintString & vbCr & vbLf
                ElseIf .nType = IO_RECT Then
                    j = GAP_X + CInt(.x * DPI_RATIO)
                    k = GAP_Y + CInt(.y * DPI_RATIO)

                    MakeCleverTagPrintString = MakeCleverTagPrintString & "BOX " & CStr(j) & ", " & CStr(k) & ", " & CStr(j + CInt(.nFigureWidth * DPI_RATIO)) & ", " & CStr(k + CInt(.nFigureHeight * DPI_RATIO)) & ", " & CStr(.nThickness) & vbCr & vbLf
                End If
            End If
        End With
    Next i

    ' Make Print Defect Fields *********************************************************************
    With m_tTag
        If .nClss > 0 And .nDefHeight > 0 And nDefectCnt > 0 Then
            If .nClss = 1 Then      ' Free Style
                Dim nCount%

                y = .nDefBaseY
                x = .nDefBaseX1
                j = .nDefBaseX2

                nCount = 1
                For i = 0 To UBound(vDefect, 1)
                    MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & _
                        CStr(GAP_X + CInt(x * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(y * DPI_RATIO)) & ", " & _
                        Chr(34) & "3" & Chr(34) & ", " & "0, 1, 1, " & Chr(34) & CStr(vDefect(i, 1)) & Chr(34) & vbCr & vbLf
                    MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & _
                        CStr(GAP_X + CInt(j * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(y * DPI_RATIO)) & ", " & _
                        Chr(34) & "3" & Chr(34) & ", " & "0, 1, 1, " & Chr(34) & Trim(vDefect(i, 0)) & CStr(vDefect(i, 2)) & Chr(34) & vbCr & vbLf

                    y = y + .nDefGapY
                    If nCount >= .nDefHCount Then
                        y = .nDefBaseY
                        x = x + .nDefGapX1
                        j = j + .nDefGapX2
                        nCount = 1
                    Else
                        nCount = nCount + 1
                    End If
                Next i
            ElseIf .nClss = 2 Then  ' Fixed Style
                For i = 0 To UBound(vDefect, 1)
                    j = vDefect(i, 1)
                    y = .nDefBaseY + ((j - 1) Mod .nDefHCount) * .nDefGapY
                    x = .nDefBaseX2 + (ArithUpper(j / .nDefHCount) - 1) * .nDefGapX2

                    MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & _
                        CStr(GAP_X + CInt(x * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(y * DPI_RATIO)) & ", " & _
                        Chr(34) & "3" & Chr(34) & ", " & "0, 1, 1, " & Chr(34) & vDefect(i, 0) & CStr(vDefect(i, 2)) & Chr(34) & vbCr & vbLf
                Next i
            End If

            If .nDefectClss > 0 Then
                 Set rs = oCode.GetCode(CD_DEFECT)

                y = .nDefBaseY
                x = .nDefBaseX3
                If .nDefectClss = 1 Then
                    Do Until rs.EOF
                        MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & _
                            CStr(GAP_X + CInt(x * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(y * DPI_RATIO)) & ", " & _
                            Chr(34) & "1" & Chr(34) & ", " & "0, 1, 1, " & Chr(34) & GetDefectText(rs!TagName, rs!EDefect, .nDefLength) & Chr(34) & vbCr & vbLf

                        y = y + 20
                        rs.MoveNext
                    Loop
                ElseIf .nDefectClss = 2 Then
                    Do Until rs.EOF
                        For i = 0 To UBound(vDefect, 1)
                            If Trim(rs!TagName) = Trim(vDefect(i, 0)) Then Exit For
                        Next i

                        If i < UBound(vDefect, 1) Then
                            MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & _
                                CStr(GAP_X + CInt(x * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(y * DPI_RATIO)) & ", " & _
                                Chr(34) & "1" & Chr(34) & ", " & "0, 1, 1, " & Chr(34) & GetDefectText(rs!TagName, rs!EDefect, .nDefLength) & Chr(34) & vbCr & vbLf

                            y = y + 20
                        End If

                        rs.MoveNext
                    Loop
                End If
            End If
        End If
    End With

    ' Make Print Defect Bar ************************************************************************
    With m_tTag
        If .nDefBarClss > 0 Then
            Dim nWidth%, nHalfWidth%, nHeight%, nThick%, nPos#

            x = .nWidth - 100
            y = 100
            nWidth = 20
            nHalfWidth = nWidth / 2
            nHeight = (.nHeight + .nDefHeight - y) - 200
            nThick = 3

            MakeCleverTagPrintString = MakeCleverTagPrintString & "BAR " & _
                CStr(GAP_X + CInt(x * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(y * DPI_RATIO)) & ", " & _
                CStr(CInt(nThick * DPI_RATIO)) & ", " & CStr(CInt(nHeight * DPI_RATIO)) & ", " & vbCr & vbLf
            MakeCleverTagPrintString = MakeCleverTagPrintString & "BAR " & _
                CStr(GAP_X + CInt((x - nHalfWidth) * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(y * DPI_RATIO)) & ", " & _
                CStr(CInt(nWidth * DPI_RATIO)) & ", " & CStr(CInt(nThick * DPI_RATIO)) & ", " & vbCr & vbLf
            MakeCleverTagPrintString = MakeCleverTagPrintString & "BAR " & _
                CStr(GAP_X + CInt((x - nHalfWidth) * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt((y + nHeight) * DPI_RATIO)) & ", " & _
                CStr(CInt(nWidth * DPI_RATIO)) & ", " & CStr(CInt(nThick * DPI_RATIO)) & ", " & vbCr & vbLf

            For i = 0 To UBound(vDefect, 1)
                nPos = y + ((vDefect(i).Position / nTotalLength) * nHeight)
                MakeCleverTagPrintString = MakeCleverTagPrintString & "BAR " & _
                    CStr(GAP_X + CInt((x - nHalfWidth) * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(nPos * DPI_RATIO)) & ", " & _
                    CStr(CInt(nWidth * DPI_RATIO)) & ", " & CStr(CInt(nThick * DPI_RATIO)) & ", " & vbCr & vbLf
                MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & _
                    CStr(GAP_X + CInt((x - nHalfWidth - 27) * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt((nPos - 5) * DPI_RATIO)) & ", " & _
                    Chr(34) & "1" & Chr(34) & ", " & "0, 1, 1, " & Chr(34) & CStr(vDefect(i, 1)) & Chr(34) & vbCr & vbLf
                MakeCleverTagPrintString = MakeCleverTagPrintString & "TEXT " & _
                    CStr(GAP_X + CInt((x + nHalfWidth + 5) * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt((nPos - 5) * DPI_RATIO)) & ", " & _
                    Chr(34) & "1" & Chr(34) & ", " & "0, 1, 1, " & Chr(34) & vDefect(i, 0) & Chr(34) & vbCr & vbLf
            Next i
        End If
    End With

    ' Make Print Tail ******************************************************************************
    MakeCleverTagPrintString = MakeCleverTagPrintString & "PRINT 1, " & CStr(nPrintCount) & vbCr & vbLf

    ReDim m_tItem(0)
    Set m_picPrint = Nothing

    Exit Function

ErrHandler:
    ReDim m_tItem(0)
    Set m_picPrint = Nothing

    MakeCleverTagPrintString = ""

    Set rsItem = Nothing
    Set rsTag = Nothing
    Set rs = Nothing
    Err.Raise Err.Number, "TagPrint.MakeCleverTagPrintString", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetCleverFontGapY(ByVal nSize As Integer) As Integer
    Select Case nSize
    Case 1
        GetCleverFontGapY = 0
    Case 2
        GetCleverFontGapY = 7
    Case 3
        GetCleverFontGapY = 9
    Case 4
        GetCleverFontGapY = 13
    Case 5
        GetCleverFontGapY = 18
    End Select
End Function

Private Function GetCleverFontDot(ByVal nSize As Integer) As Integer
    Select Case nSize
    Case 1
        GetCleverFontDot = 28
    Case 2
        GetCleverFontDot = 40
    Case 3
        GetCleverFontDot = 48
    Case 4
        GetCleverFontDot = 64
    Case 5
        GetCleverFontDot = 94
    End Select
End Function

Private Function GetCleverHangulFontSize(nSize As Integer) As Integer
    Select Case nSize
    Case 1
        GetCleverHangulFontSize = 15
    Case 2
        GetCleverHangulFontSize = 24
    Case 3
        GetCleverHangulFontSize = 30
    Case 4
        GetCleverHangulFontSize = 40
    Case 5
        GetCleverHangulFontSize = 60
    End Select
End Function

Public Function MakeZebraTagPrintString(vData As Variant, ByVal sTagID As String, Optional vDefect As Variant, Optional nTotalLength As Integer, Optional m_nShipCnt As Integer) As String
    Dim oTag   As PlusLib2.CTag
    Dim oCode  As PlusLib2.CCode
    Dim rs     As ADODB.Recordset
    Dim rsTag  As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    Dim i&, j&, k&, l&, nLen&, x&, y&, nPixWidth%, nPixHeight%, nPixel&, iImage%
    Dim sText$, sChar$, nBitmap As Byte

    On Error GoTo ErrHandler

    ' Data Loads ***********************************************************************************
    Set oTag = New PlusLib2.CTag
    oTag.Connection = g_adoCon
    
    Set rsTag = oTag.GetTag(1, sTagID)
    Set rsItem = oTag.GetTagSub(sTagID)


    With m_tTag
        .sTag = rsTag!Tag
        .nWidth = rsTag!Width
        .nHeight = rsTag!Height
        .nClss = rsTag!Clss
        .nDefectClss = rsTag!DefectClss
        .nDefHeight = rsTag!DefHeight
        .nDefBaseY = rsTag!DefBaseY
        .nDefBaseX1 = rsTag!DefBaseX1
        .nDefBaseX2 = rsTag!DefBaseX2
        .nDefBaseX3 = rsTag!DefBaseX3
        .nDefGapY = rsTag!DefGapY
        .nDefGapX1 = rsTag!DefGapX1
        .nDefGapX2 = rsTag!DefGapX2
        .nDefLength = rsTag!DefLength
        .nDefHCount = rsTag!DefHCount
        .nDefBarClss = rsTag!DefBarClss
        .nGap = rsTag!Gap
    End With
    rsTag.Close
    Set rsTag = Nothing

    Do Until rsItem.EOF
        ReDim Preserve m_tItem(i)

        With m_tItem(i)
            .sName = rsItem!Name
            .nType = rsItem!Type
            .nAlign = rsItem!Align
            .x = rsItem!x
            .y = rsItem!y
            .nFont = rsItem!Font
            .nLength = rsItem!Length
            .nHMulti = rsItem!HMulti
            .nVMulti = rsItem!VMulti
            .nRotation = rsItem!Rotation
            .sText = rsItem!Text
            .nSpace = rsItem!Space
            .nRelation = rsItem!Relation
            .nPrevItem = rsItem!PrevItem
            .nBarType = rsItem!BarType
            .nBarHeight = rsItem!BarHeight
            .nFigureWidth = rsItem!FigureWidth
            .nFigureHeight = rsItem!FigureHeight
            .nThickness = rsItem!Thickness
            .sImageFile = rsItem!ImageFile
            .nWidth = rsItem!Width
            .nHeight = rsItem!Height
            .nVisible = rsItem!Visible
        End With

        i = i + 1
        rsItem.MoveNext
    Loop
    rsItem.Close
    Set rsItem = Nothing

    ' Make Print Head ******************************************************************************
    MakeZebraTagPrintString = "^XA" & vbCr & vbLf & _
        "^LH50,0" & vbCr & vbLf & _
        "^PO" & IIf(CInt(m_tTag.sDirect) = 0, "N", "I") & vbCr & vbLf

    Set m_picPrint = frmInspect.picPrint
    Set m_picFont = frmInspect.picPrint

    ' Make Print Fields ****************************************************************************
    ReDim m_sData(UBound(vData))
    For i = 0 To UBound(vData)
        m_sData(i) = vData(i)
    Next i

    For i = 0 To UBound(m_tItem)
        With m_tItem(i)
            If .nVisible > 0 Then
                If .nType = IO_BARCODE Then
                    If .nPrevItem = 0 Then
                        If .nRotation = 0 Then  ' 0 Angle
                            MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(.x * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & _
                                "^B3N,N," & CStr(CInt(.nBarHeight * DPI_RATIO)) & ",N^FD" & GetBarCodeItemText(i) & "^FS" & vbCr & vbLf
                        Else
                            MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt((.x + .nBarHeight) * DPI_RATIO)) & ", " & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & _
                                "^B3R,N," & CStr(CInt(.nBarHeight * DPI_RATIO)) & ",N^FD" & GetBarCodeItemText(i) & "^FS" & vbCr & vbLf
                        End If
                    End If
                ElseIf .nType = IO_DATA Or .nType = IO_TEXT Then
                    Dim nFontHeight%, nFontWidth%, nGapY%, nHMulti%, nVMulti%

                    sText = GetItemText(i)
                    nHMulti = .nHMulti + 1
                    nVMulti = .nVMulti + 1
                    nFontHeight = GetZebraFontHeight(.nFont) * nVMulti
                    nFontWidth = GetZebraFontWidth(.nFont) * nHMulti
                    nGapY = GetZebraFontGapY(.nFont) / nVMulti

                    If IsHangul(sText) Then
                        If .nRotation = 0 Then  ' 0 Angle
                            x = GAP_X + CInt(.x * DPI_RATIO)
                            nLen = Len(sText) - 1
                            For j = 0 To nLen
                                sChar = Mid(sText, j + 1, 1)
                                If IsHangul(sChar) Then
                                    MakeZebraTagPrintString = MakeZebraTagPrintString & "~DGR:IMAGE" & Format(iImage, "000") & ".GRF," & _
                                        ArithUpper(CSng(nFontHeight * nVMulti * DPI_RATIO)) * ArithUpper(CSng((nFontHeight * nHMulti * DPI_RATIO) / 8)) & "," & ArithUpper(CSng((nFontHeight * nHMulti * DPI_RATIO) / 8)) & "," & _
                                        GetHangulBitmap(1, .nFont, sChar, .nRotation, nHMulti, nVMulti)
                                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(x) & "," & CStr(GAP_Y + CInt(.y * DPI_RATIO) - 2) & _
                                        "^XGR:IMAGE" & Format(iImage, "000") & ".GRF,1,1^FS" & vbCr & vbLf
                                    iImage = iImage + 1

                                    x = x + (nFontHeight * 0.9) * nHMulti
                                Else
                                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(x) & "," & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & _
                                        "^FWN^A0," & CStr(nFontHeight) & "," & CStr(nFontWidth) & "^FD" & sChar & "^FS" & vbCr & vbLf

                                    x = x + (nFontWidth * (m_picFont.TextWidth(sChar) / m_picFont.TextWidth("A")) * 0.55)
                                End If
                            Next j
                        Else                    ' 90 Angle
                            y = GAP_Y + CInt(.y * DPI_RATIO)
                            nLen = Len(sText) - 1
                            For j = 0 To nLen
                                sChar = Mid(sText, j + 1, 1)
                                If IsHangul(sChar) Then
                                    MakeZebraTagPrintString = MakeZebraTagPrintString & "~DGR:IMAGE" & Format(iImage, "000") & ".GRF," & _
                                        ArithUpper(CSng(nFontHeight * nHMulti * DPI_RATIO)) * ArithUpper(CSng((nFontHeight * nVMulti * DPI_RATIO) / 8)) & "," & ArithUpper(CSng((nFontHeight * nVMulti * DPI_RATIO) / 8)) & "," & _
                                        GetHangulBitmap(1, .nFont, sChar, .nRotation, nHMulti, nVMulti)
                                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(.x * DPI_RATIO) + GetZebraFontGapY(.nFont)) & "," & CStr(y) & _
                                        "^XGR:IMAGE" & Format(iImage, "000") & ".GRF,1,1^FS" & vbCr & vbLf
                                    iImage = iImage + 1

                                    y = y + (nFontHeight * 0.9) * nHMulti
                                Else
                                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(.x * DPI_RATIO)) & "," & CStr(y) & _
                                        "^FWR^A0," & CStr(nFontHeight) & "," & CStr(nFontWidth) & "^FD" & sChar & "^FS" & vbCr & vbLf

                                    y = y + (nFontWidth * (m_picFont.TextWidth(sChar) / m_picFont.TextWidth("A")) * 0.55)
                                End If
                            Next j
                        End If
                    Else
                        If .nRotation = 0 Then  ' 0 Angle
                            MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(.x * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & _
                                "^FWN^A0," & CStr(nFontHeight) & "," & CStr(nFontWidth) & "^FD" & sText & "^FS" & vbCr & vbLf
                        Else
                            MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(.x * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & _
                                "^FWR^A0," & CStr(nFontHeight) & "," & CStr(nFontWidth) & "^FD" & sText & "^FS" & vbCr & vbLf
                        End If
                    End If
                ElseIf .nType = IO_LINE And (.nFigureHeight <= 5 Or .nFigureWidth <= 5) Then
                    j = GAP_X + CInt(.x * DPI_RATIO)
                    k = GAP_Y + CInt(.y * DPI_RATIO)

                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(j) & "," & CStr(k) & "^GB" & CStr(CInt(.nFigureWidth * DPI_RATIO)) & "," & CStr(CInt(.nFigureHeight * DPI_RATIO)) & "," & CStr(.nThickness) & "^FS" & vbCr & vbLf
                ElseIf .nType = IO_LINE Or .nType = IO_DIAMOND Or .nType = IO_IMAGE Then
                    Dim nStart%, nVWidth%, nVHeight%, nAdjust#

                    m_picPrint.Width = ArithUpper(.nFigureWidth * DPI_RATIO) * Screen.TwipsPerPixelX + (Screen.TwipsPerPixelX * 2)
                    m_picPrint.Height = ArithUpper(.nFigureHeight * DPI_RATIO) * Screen.TwipsPerPixelY + (Screen.TwipsPerPixelY * 2)
                    m_picPrint.ScaleWidth = m_picPrint.Width - Screen.TwipsPerPixelX * 2
                    m_picPrint.ScaleHeight = m_picPrint.Height - Screen.TwipsPerPixelY * 2

                    m_picPrint.Picture = Nothing
                    m_picPrint.Cls

                    nAdjust = .nThickness / 2
                    nStart = Screen.TwipsPerPixelX * ArithLower(nAdjust)
                    nVWidth = m_picPrint.ScaleWidth - Screen.TwipsPerPixelX * ArithUpper(nAdjust)
                    nVHeight = m_picPrint.ScaleHeight - Screen.TwipsPerPixelY * ArithUpper(nAdjust)

                    m_picPrint.DrawWidth = .nThickness
                    If .nType = IO_LINE Then
                        If .nAlign = 0 Then ' like \
                            m_picPrint.Line (0, 0)-(m_picPrint.ScaleWidth, m_picPrint.ScaleHeight)
                        Else                ' like /
                            m_picPrint.Line (m_picPrint.ScaleWidth, 0)-(0, m_picPrint.ScaleHeight)
                        End If
                    ElseIf .nType = IO_RECT Then
                        m_picPrint.Line (nStart, nStart)-(nVWidth, nStart)
                        m_picPrint.Line -(nVWidth, nVHeight)
                        m_picPrint.Line -(nStart, nVHeight)
                        m_picPrint.Line -(nStart, nStart)
                    ElseIf .nType = IO_DIAMOND Then
                        m_picPrint.Line (nVWidth / 2, nStart)-(nVWidth, nVHeight / 2)
                        m_picPrint.Line -(nVWidth / 2, nVHeight)
                        m_picPrint.Line -(nStart, nVHeight / 2)
                        m_picPrint.Line -(nVWidth / 2, nStart)
                    ElseIf .nType = IO_CIRCLE Then
                        m_picPrint.Circle (ArithLower(m_picPrint.ScaleWidth / 2), ArithLower(m_picPrint.ScaleHeight / 2)), _
                            ArithLower(IIf(m_picPrint.ScaleWidth > m_picPrint.ScaleHeight, nVWidth, nVHeight) / 2 - Screen.TwipsPerPixelX), , , , _
                            CDbl(m_picPrint.ScaleHeight / m_picPrint.ScaleWidth)
                    ElseIf .nType = IO_IMAGE Then
                        m_picPrint.Picture = LoadPicture(.sImageFile)
                    End If

                    MakeZebraTagPrintString = MakeZebraTagPrintString & "~DGR:IMAGE" & Format(iImage, "000") & ".GRF," & _
                        ArithUpper(CSng(.nFigureHeight * DPI_RATIO)) * ArithUpper(CSng((.nFigureWidth * DPI_RATIO) / 8)) & "," & ArithUpper(CSng((.nFigureWidth * DPI_RATIO) / 8)) & ","

                    nPixWidth = m_picPrint.ScaleWidth / Screen.TwipsPerPixelX - 1
                    nPixHeight = m_picPrint.ScaleHeight / Screen.TwipsPerPixelY - 1
                    For j = 0 To nPixHeight
                        nBitmap = 0
                        For k = 0 To nPixWidth
                            l = k Mod 8

                            nPixel = GetPixel(m_picPrint.hdc, k, j)
                            If nPixel = 0 Then
                                nBitmap = nBitmap + (2 ^ Abs(l - 7))
                            End If

                            If l = 7 Or k = nPixWidth Then
                                MakeZebraTagPrintString = MakeZebraTagPrintString & BYTEtoZebraSTR(nBitmap)
                                nBitmap = 0
                            End If
                        Next k
                    Next j
                    MakeZebraTagPrintString = MakeZebraTagPrintString & vbCr & vbLf

                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(.x * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(.y * DPI_RATIO)) & _
                        "^XGR:IMAGE" & Format(iImage, "000") & ".GRF,1,1^FS" & vbCr & vbLf

                    iImage = iImage + 1
                ElseIf .nType = IO_RECT Then
                    j = GAP_X + CInt(.x * DPI_RATIO)
                    k = GAP_Y + CInt(.y * DPI_RATIO)

                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(j) & "," & CStr(k) & "^GB" & CStr(CInt(.nFigureWidth * DPI_RATIO)) & "," & CStr(CInt(.nFigureHeight * DPI_RATIO)) & "," & CStr(.nThickness) & "^FS" & vbCr & vbLf
                End If
            End If
        End With
    Next i

    ' Make Print Defect Fields *********************************************************************
    With m_tTag
        If .nClss > 0 And .nDefHeight > 0 Then
            If .nClss = 1 Then      ' Free Style
                Dim nCount%

                y = .nDefBaseY
                x = .nDefBaseX1
                j = .nDefBaseX2

                nCount = 1
                For i = 0 To UBound(vDefect, 1)
                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(x * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(y * DPI_RATIO)) & _
                        "^FWN^A0,18,16^FD" & CStr(vDefect(i, 1)) & "^FS" & vbCr & vbLf
                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(j * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(y * DPI_RATIO)) & _
                        "^FWN^A0,18,16^FD" & vDefect(i, 0) & "^FS" & vbCr & vbLf

                    y = y + .nDefGapY
                    If nCount >= .nDefHCount Then
                        y = .nDefBaseY
                        x = x + .nDefGapX1
                        j = j + .nDefGapX2
                        nCount = 1
                    Else
                        nCount = nCount + 1
                    End If
                Next i
            ElseIf .nClss = 2 Then  ' Fixed Style
                For i = 0 To UBound(vDefect, 1)
                    j = vDefect(i, 1)
                    y = .nDefBaseY + ((j - 1) Mod .nDefHCount) * .nDefGapY
                    x = .nDefBaseX2 + (ArithUpper(j / .nDefHCount) - 1) * .nDefGapX2

                    MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(x * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(y * DPI_RATIO)) & _
                        "^FWN^A0,18,16^FD" & Trim(vDefect(i, 0)) & CStr(vDefect(i, 2)) & "^FS" & vbCr & vbLf
                Next i
            End If

            If .nDefectClss > 0 Then
                Set rs = oCode.GetCode(CD_DEFECT)

                y = .nDefBaseY
                x = .nDefBaseX3
                If .nDefectClss = 1 Then
                    Do Until rs.EOF
                        MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(x * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(y * DPI_RATIO)) & _
                            "^FWN^A0,18,16^FD" & GetDefectText(rs!TagName, rs!EDefect, .nDefLength) & "^FS" & vbCr & vbLf

                        y = y + 20
                        rs.MoveNext
                    Loop
                ElseIf .nDefectClss = 2 Then
                    Do Until rs.EOF
                        For i = 0 To UBound(vDefect, 1)
                            If Trim(rs!TagName) = Trim(vDefect(i, 0)) Then Exit For
                        Next i

                        If i < UBound(vDefect, 1) Then
                            MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(x * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(y * DPI_RATIO)) & _
                                "^FWN^A0,18,16^FD" & GetDefectText(rs!TagName, rs!EDefect, .nDefLength) & "^FS" & vbCr & vbLf

                            y = y + 20
                        End If

                        rs.MoveNext
                    Loop
                End If
            End If
        End If
    End With

    ' Make Print Defect Bar ************************************************************************
    With m_tTag
        If .nDefBarClss > 0 Then
            Dim nWidth%, nHalfWidth%, nHeight%, nThick%, nPos#

            x = .nWidth - 100
            y = 100
            nWidth = 20
            nHalfWidth = nWidth / 2
            nHeight = (.nHeight + .nDefHeight - y) - 200
            nThick = 3

            MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt(x * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(y * DPI_RATIO)) & _
                "^GB0," & CStr(CInt(nHeight * DPI_RATIO)) & "," & CStr(CInt(nThick * DPI_RATIO)) & "^FS" & vbCr & vbLf
            MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt((x - nHalfWidth) * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(y * DPI_RATIO)) & _
                "^GB" & CInt(nWidth * DPI_RATIO) & ",0," & CStr(CInt(nThick * DPI_RATIO)) & "^FS" & vbCr & vbLf
            MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt((x - nHalfWidth) * DPI_RATIO)) & "," & CStr(GAP_Y + CInt((y + nHeight) * DPI_RATIO)) & _
                "^GB" & CInt(nWidth * DPI_RATIO) & ",0," & CStr(CInt(nThick * DPI_RATIO)) & "^FS" & vbCr & vbLf

            For i = 0 To UBound(vDefect, 1)
                nPos = y + ((vDefect(i, 1) / nTotalLength) * nHeight)

                MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt((x - nHalfWidth) * DPI_RATIO)) & "," & CStr(GAP_Y + CInt(nPos * DPI_RATIO)) & _
                    "^GB" & CInt(nWidth * DPI_RATIO) & ",0," & CStr(CInt(nThick * DPI_RATIO)) & "^FS" & vbCr & vbLf
                MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt((x - nHalfWidth - 27) * DPI_RATIO)) & "," & CStr(GAP_Y + CInt((nPos - 5) * DPI_RATIO)) & _
                    "^FWN^A0,18,16^FD" & CStr(vDefect(i, 1)) & "^FS" & vbCr & vbLf
                MakeZebraTagPrintString = MakeZebraTagPrintString & "^FO" & CStr(GAP_X + CInt((x + nHalfWidth + 5) * DPI_RATIO)) & "," & CStr(GAP_Y + CInt((nPos - 5) * DPI_RATIO)) & _
                    "^FWN^A0,18,16^FD" & vDefect(i, 0) & "^FS" & vbCr & vbLf
            Next i
        End If
    End With

    MakeZebraTagPrintString = MakeZebraTagPrintString & "^PQ" & CStr(m_nShipCnt) & vbCr & vbLf
    MakeZebraTagPrintString = MakeZebraTagPrintString & "^XZ"

    Debug.Print MakeZebraTagPrintString

    ReDim m_tItem(0)
    Set m_picFont = Nothing
    Set m_picPrint = Nothing

    Exit Function

ErrHandler:
    ReDim m_tItem(0)
    Set oTag = Nothing
    Set m_picFont = Nothing
    Set m_picPrint = Nothing

    MakeZebraTagPrintString = ""

    Set rsItem = Nothing
    Set rsTag = Nothing
    Set rs = Nothing
    Set oTag = Nothing
    Err.Raise Err.Number, "TagPrint.MakeZebraTagPrintString", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetZebraFontGapY(ByVal nSize As Integer) As Integer
    Select Case nSize
    Case 1
        GetZebraFontGapY = 2
    Case 2
        GetZebraFontGapY = 5
    Case 3
        GetZebraFontGapY = 6
    Case 4
        GetZebraFontGapY = 10
    Case 5
        GetZebraFontGapY = 15
    End Select
End Function

Private Function GetZebraFontHeight(ByVal nSize As Integer) As Integer
    Select Case nSize
    Case 1
        GetZebraFontHeight = 18
    Case 2
        GetZebraFontHeight = 27
    Case 3
        GetZebraFontHeight = 34
    Case 4
        GetZebraFontHeight = 50
    Case 5
        GetZebraFontHeight = 73
    End Select
End Function

Private Function GetZebraFontWidth(ByVal nSize As Integer) As Integer
    Select Case nSize
    Case 1
        GetZebraFontWidth = 16
    Case 2
        GetZebraFontWidth = 23
    Case 3
        GetZebraFontWidth = 32
    Case 4
        GetZebraFontWidth = 44
    Case 5
        GetZebraFontWidth = 67
    End Select
End Function

Private Function GetZebraHangulFontSize(nSize As Integer) As Integer
    Select Case nSize
    Case 1
        GetZebraHangulFontSize = 11
    Case 2
        GetZebraHangulFontSize = 17
    Case 3
        GetZebraHangulFontSize = 22
    Case 4
        GetZebraHangulFontSize = 31
    Case 5
        GetZebraHangulFontSize = 46
    End Select
End Function

Private Function BYTEtoZebraSTR(nData As Byte) As String
    Dim nLeft%, nRight%

    nRight = nData Mod 16
    nLeft = ArithLower(nData / 16)

    BYTEtoZebraSTR = Chr(IIf(nLeft < 10, 48 + nLeft, 55 + nLeft)) & Chr(IIf(nRight < 10, 48 + nRight, 55 + nRight))
End Function

Private Function GetHangulBitmap(ByVal nPrinter As Integer, ByVal nFont As Integer, ByVal sText As String, ByVal nRotation As Integer, Optional ByVal nHMulti As Integer = 0, Optional ByVal nVMulti As Integer = 0) As String
    Dim i&, j&, k&, nPixWidth%, nPixHeight%, nPixel&, nBitmap As Byte

    With m_picPrint
        If nPrinter = 0 Then    ' Clever
            nPixWidth = GetCleverFontDot(nFont)
            .FontSize = GetCleverHangulFontSize(nFont)
        Else
            nPixWidth = GetZebraFontHeight(nFont)
            .FontSize = GetZebraHangulFontSize(nFont)
        End If
        .FontName = FONT_NAME
        .FontBold = True
        .BackColor = &HFFFFFF

        .Width = ArithUpper(nPixWidth * DPI_RATIO) * Screen.TwipsPerPixelX + (Screen.TwipsPerPixelX * 2)
        .Height = .Width
        .ScaleWidth = .Width - Screen.TwipsPerPixelX * 2
        .ScaleHeight = .ScaleWidth

        .Cls

        .CurrentX = 0
        .CurrentY = 0
        m_picPrint.Print sText

        nPixWidth = .ScaleWidth

        If nHMulti > 1 Or nVMulti > 1 Then
            .Width = .Width * nHMulti - (Screen.TwipsPerPixelX * 2) * (nHMulti - 1)
            .Height = .Height * nVMulti - (Screen.TwipsPerPixelY * 2) * (nVMulti - 1)
            .ScaleWidth = .Width - Screen.TwipsPerPixelX * 2
            .ScaleHeight = .Height - Screen.TwipsPerPixelY * 2

            Call StretchBlt(.hdc, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, _
                .hdc, 0, 0, nPixWidth / Screen.TwipsPerPixelX, nPixWidth / Screen.TwipsPerPixelY, vbSrcCopy)
        End If

        If nRotation = 1 Then
            Dim lPixel() As Long

            nPixWidth = .ScaleWidth / Screen.TwipsPerPixelX
            nPixHeight = .ScaleHeight / Screen.TwipsPerPixelY

            ReDim lPixel(nPixWidth - 1, nPixHeight - 1) As Long
            For i = 0 To nPixWidth - 1
                For j = 0 To nPixHeight - 1
                    lPixel(i, j) = GetPixel(.hdc, i, j)
                Next j
            Next i

            i = .Width
            j = .ScaleWidth
            .Width = .Height
            .ScaleWidth = .ScaleHeight
            .Height = i
            .ScaleHeight = j

            .Cls

            Select Case nRotation
            Case 1
                For i = 0 To nPixWidth - 1
                    For j = 0 To nPixHeight - 1
                        Call SetPixel(.hdc, nPixHeight - j, i, lPixel(i, j))
                    Next j
                Next i
            End Select
        End If

        nPixWidth = .ScaleWidth / Screen.TwipsPerPixelX - 1
        nPixHeight = .ScaleHeight / Screen.TwipsPerPixelY - 1

        If nPrinter = 0 Then    ' Clever
            For i = 0 To nPixHeight
                nBitmap = 255
                For j = 0 To nPixWidth
                    k = j Mod 8

                    nPixel = GetPixel(m_picPrint.hdc, j, i)
                    If nPixel = 0 Then nBitmap = nBitmap - (2 ^ Abs(k - 7))

                    If k = 7 Or j = nPixWidth Then
                        GetHangulBitmap = GetHangulBitmap & ChrB(nBitmap) & ChrB(0)
                        nBitmap = 255
                    End If
                Next j
            Next i
        ElseIf nPrinter = 1 Then    ' Zebra
            For i = 0 To nPixHeight
                nBitmap = 0
                For j = 0 To nPixWidth
                    k = j Mod 8

                    nPixel = GetPixel(m_picPrint.hdc, j, i)
                    If nPixel = 0 Then nBitmap = nBitmap + (2 ^ Abs(k - 7))

                    If k = 7 Or j = nPixWidth Then
                        GetHangulBitmap = GetHangulBitmap & BYTEtoZebraSTR(nBitmap)
                        nBitmap = 0
                    End If
                Next j
            Next i
        End If
    End With
End Function

Private Function GetDefectText(ByVal sTag As String, ByVal sDefect As String, ByVal nLen As Integer) As String
    Dim i%, nDefectLen%, nCount%

    sDefect = Trim(sDefect)
    nDefectLen = LenB(StrConv(sDefect, vbFromUnicode))

    If nDefectLen > nLen - 2 Then
        sDefect = LeftB(StrConv(sDefect, vbFromUnicode), nLen - 2)
        nCount = 1
    Else
        nCount = nLen - nDefectLen - 1
    End If

    GetDefectText = Left(Trim(sTag), 1)
    For i = 1 To nCount
        GetDefectText = GetDefectText & "-"
    Next i
    GetDefectText = GetDefectText & sDefect
End Function

Private Function GetBarCodeItemText(ByVal iItem As Integer) As String
    Dim i%, bFound As Boolean

    With m_tItem(iItem)
        GetBarCodeItemText = GetItemText(iItem)

        Do
            bFound = False

            For i = 0 To UBound(m_tItem)
                If m_tItem(i).nPrevItem - 1 = iItem Then
                    GetBarCodeItemText = GetBarCodeItemText & GetItemText(i)
                    iItem = i
                    bFound = True
                    Exit For
                End If
            Next i
        Loop While bFound
    End With
End Function

Private Function GetItemText(ByVal iItem As Integer) As String
    Dim i%, nLen%, sText$, sChar$
    Dim iIdx%, sTempChar$, sTempText$

    On Error Resume Next

    With m_tItem(iItem)
        sText = Trim(m_sData(.nRelation))
        nLen = LenB(StrConv(sText, vbFromUnicode))

        If .nSpace = 0 Then
            sChar = " "
        ElseIf .nSpace = 1 Then
            sChar = "0"
        End If

        If .nAlign = 0 Then ' Align Left
            If nLen > .nLength Then
                iIdx = 1
                sTempText = ""
                For i = 1 To Abs(.nLength)
                    sTempChar = Mid(sText, iIdx, 1)
                    iIdx = iIdx + 1
                    sTempText = sTempText & sTempChar
                    If IsHangul(sTempChar) Then i = i + 1
                Next i
                sText = sTempText
            ElseIf nLen < .nLength Then
                For i = nLen To .nLength - 1
                    sText = sText & sChar
                Next i
            End If
        Else                ' Align Right
            If nLen > .nLength Then
                iIdx = Len(sText)
                sTempText = ""
                For i = 1 To Abs(.nLength)
                    sTempChar = Mid(sText, iIdx, 1)
                    iIdx = iIdx - 1
                    sTempText = sTempChar & sTempText
                    If IsHangul(sTempChar) Then i = i + 1
                Next i
                sText = sTempText
            ElseIf nLen < .nLength Then
                For i = nLen To .nLength - 1
                    sText = sChar & sText
                Next i
            End If
        End If

        GetItemText = sText
    End With
End Function

Private Function IsHangul(sText As String) As Boolean
    IsHangul = IIf(Len(sText) = LenB(StrConv(sText, vbFromUnicode)), False, True)
End Function

Public Function ArithUpper(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double
    x = x + (0.49 * Factor)
    ArithUpper = Fix(x * Factor + 0.5 * Sgn(x)) / Factor
End Function

Public Function ArithLower(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double
    x = x - (0.49 * Factor)
    ArithLower = Fix(x * Factor + 0.5 * Sgn(x)) / Factor
End Function


