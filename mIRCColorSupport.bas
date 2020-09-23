Attribute VB_Name = "mIRCColorSupport"

Option Explicit

'Control code numbers
Public Const SP_BOLD    As Integer = 2
Public Const SP_COLOR   As Integer = 3
Public Const SP_NORMAL  As Integer = 15
Public Const SP_INVERSE As Integer = 22
Public Const SP_ITALIC  As Integer = 29
Public Const SP_UNDER   As Integer = 31

'Control code characters (faster than doing ChrW$()
Public Const CC_BOLD    As String * 1 = ""
Public Const CC_COLOR   As String * 1 = ""
Public Const CC_NORMAL  As String * 1 = ""
Public Const CC_INVERSE As String * 1 = ""
Public Const CC_ITALIC  As String * 1 = ""    'ChrW$(29)
Public Const CC_UNDER   As String * 1 = ""
Public Const CC_1       As String * 1 = ""    'ChrW$(1)
Public Const CC_15      As String * 1 = ""  'ChrW$(15)

Public EventColor(22)   As Byte
Public IRCColor(15)     As Long

Public Sub LoadScheme()
    On Error Resume Next
    'Set event colors
    EventColor(0) = 0
    EventColor(1) = 6
    EventColor(2) = 5
    EventColor(3) = 2
    EventColor(4) = 3
    EventColor(5) = 3
    EventColor(6) = 3
    EventColor(7) = 4
    EventColor(8) = 2
    EventColor(9) = 3
    EventColor(10) = 4
    EventColor(11) = 5
    EventColor(12) = 7
    EventColor(13) = 1
    EventColor(14) = 6
    EventColor(15) = 1
    EventColor(16) = 3
    EventColor(17) = 2
    EventColor(18) = 2
    EventColor(19) = 0
    EventColor(20) = 1
    EventColor(21) = 0
    EventColor(22) = 1
    'Set RGB colors
    IRCColor(0) = 16777215
    IRCColor(1) = 0
    IRCColor(2) = 8323072
    IRCColor(3) = 37632
    IRCColor(4) = 255
    IRCColor(5) = 127
    IRCColor(6) = 10223772
    IRCColor(7) = 32764
    IRCColor(8) = 65535
    IRCColor(9) = 64512
    IRCColor(10) = 9671424
    IRCColor(11) = 16776960
    IRCColor(12) = 16515072
    IRCColor(13) = 16711935
    IRCColor(14) = 8355711
    IRCColor(15) = 13816530
End Sub

Public Function GetTimeStamp() As Byte()
    On Error Resume Next
    Dim tStamp As String, i As Long, t As Long, strTemp As String, tmpArr() As String, tmpSt As String
    'This can be used for selecting any format for our timestamp
    tStamp = "[h:nnt]"
    'This is our search string
    strTemp = "HH,H,hh,h,nn,n,ss,s,TT,Tt,tT,T,tt,t"
    tmpArr = Split(strTemp, ChrW$(44))
    'Search for matching parts
    For t = LBound(tmpArr) To UBound(tmpArr)
        i = InStr(tStamp, tmpArr(t))
        If i <> 0 Then
            If InStr(tmpArr(t), "t") + InStr(tmpArr(t), "T") = 0 Then
                If tmpArr(t) = "H" Or tmpArr(t) = "h" Then
                    'we only want the last digit if it's below 12
                    If CLng(Format$(Time$, "HH")) < 10 Then
                        tStamp = Replace$(tStamp, tmpArr(t), Right$(GetTok(Format$(Time$, tmpArr(t) & IIf(InStr(tmpArr(t), "h"), " a/p", vbNullString)), "1", 32), 1))
                    Else
                        tStamp = Replace$(tStamp, tmpArr(t), GetTok(Format$(Time$, tmpArr(t) & IIf(InStr(tmpArr(t), "h"), " a/p", vbNullString)), "1", 32))
                    End If
                Else
                    tStamp = Replace$(tStamp, tmpArr(t), GetTok(Format$(Time$, tmpArr(t) & IIf(InStr(tmpArr(t), "h"), " a/p", vbNullString)), "1", 32))
                End If
            Else
                For i = 1 To Len(tmpArr(t))
                    If i = 1 Then
                        tmpSt = Format$(Now, IIf(Left$(tmpArr(t), 1) = "T", "A/P", "a/p"))
                    Else
                        tmpSt = tmpSt & Right$(Format$(Time$, IIf(Right$(tmpArr(t), 1) = "T", "AM/PM", "am/pm")), 1)
                    End If
                Next i
                tStamp = Replace$(tStamp, tmpArr(t), tmpSt)
            End If
        End If
    Next t
    'Return time stamp
    Erase tmpArr()
    GetTimeStamp = tStamp & CC_NORMAL & ChrW$(32)
End Function

Public Function LineSep() As String
    On Error Resume Next
    LineSep = "-"
End Function

'Strip text function
Public Function StripCC(sText As String, sType As String) As String
    On Error Resume Next
    Dim nc As Integer, i As Integer, col As Integer, slen As Integer
    Dim new_str As String, X As Integer, iArray(1 To 6) As Integer
    '
    If InStr(1, UCase$(sType), "C") <> 0 Then iArray(1) = 1
    If InStr(1, UCase$(sType), "U") <> 0 Then iArray(2) = 1
    If InStr(1, UCase$(sType), "R") <> 0 Then iArray(3) = 1
    If InStr(1, UCase$(sType), "B") <> 0 Then iArray(4) = 1
    If InStr(1, UCase$(sType), "I") <> 0 Then iArray(5) = 1
    If InStr(1, UCase$(sType), "N") <> 0 Then iArray(6) = 1
    '
    nc = 0
    i = 0
    col = 0
    X = 1
    slen = Len(sText)
    '
    Do While (slen > 0)
        If (((col And isDigit(Mid$(sText, X, 1)) And (nc < 2)) Or ((col And Mid$(sText, X, 1) = ",") And (isDigit(Mid$(sText, (X + 1), 1))) And (nc < 3)))) Then
            If iArray(1) = 1 Then
                nc = nc + 1
                If (Mid$(sText, X, 1) = ",") Then nc = 0
            End If
        Else
            col = 0
            Select Case Asc(Mid$(sText, X, 1))
                Case SP_COLOR
                    If iArray(1) = 1 Then
                        col = 1
                        nc = 0
                        GoTo Skip_Byte
                    Else
                        new_str = new_str & Mid$(sText, X, 1)
                        i = i + 1
                    End If
                Case SP_UNDER
                    If iArray(2) = 1 Then
                        GoTo Skip_Byte
                    Else
                        new_str = new_str & Mid$(sText, X, 1)
                        i = i + 1
                    End If
                Case SP_INVERSE
                    If iArray(3) = 1 Then
                        GoTo Skip_Byte
                    Else
                        new_str = new_str & Mid$(sText, X, 1)
                        i = i + 1
                    End If
                Case SP_BOLD
                    If iArray(4) = 1 Then
                        GoTo Skip_Byte
                    Else
                        new_str = new_str & Mid$(sText, X, 1)
                        i = i + 1
                    End If
                Case SP_ITALIC
                    If iArray(5) = 1 Then
                        GoTo Skip_Byte
                    Else
                        new_str = new_str & Mid$(sText, X, 1)
                        i = i + 1
                    End If
                Case SP_NORMAL
                    If iArray(6) = 1 Then
                        GoTo Skip_Byte
                    Else
                        new_str = new_str & Mid$(sText, X, 1)
                        i = i + 1
                    End If
                Case Else:
                    new_str = new_str & Mid$(sText, X, 1)
                    i = i + 1
                End Select
        End If
Skip_Byte:
        X = X + 1
        slen = slen - 1
    Loop
    StripCC = new_str
End Function

Private Function isDigit(Digit As String) As Boolean
    On Error Resume Next
    Dim X As Integer, C As String
    '
    C = Left$(Digit, 1)
    If LenB(C) = 0 Then
        isDigit = False
        Exit Function
    End If
    X = Asc(C)
    If ((X >= 48) And (X <= 57)) Then
        isDigit = True
    Else
        isDigit = False
    End If
End Function

'TileBitmap function by Carls P.V.
Public Function TileBitmap(Picture As StdPicture, ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    Dim tBI As BITMAP, tBIH As BITMAPINFOHEADER
    Dim Buff() As Byte, lHDC As Long, lhOldBmp As Long
    Dim TileRect As RECT, PtOrg As POINTAPI, m_hBrush As Long
    Dim lPicMode As Long, i As Long
    '
    lPicMode = SetStretchBltMode(hdc, STRETCH_HALFTONE)
    '
    If (GetObjectType(Picture) = 7) Then
        'Get image info
        GetObject Picture, Len(tBI), tBI
        'Prepare DIB header and redim. Buff() array
        With tBIH
            .biSize = Len(tBIH) '40
            .biPlanes = 1
            .biBitCount = 24
            .biWidth = tBI.bmWidth
            .biHeight = tBI.bmHeight
            .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        End With
        '
        ReDim Buff(1 To Len(tBIH) + tBIH.biSizeImage) '[Header + Bits]
        'Create DIB brush
        lHDC = CreateCompatibleDC(0)
        If (lHDC <> 0) Then
            lhOldBmp = SelectObject(lHDC, Picture)
            'Build packed DIB:
            'Merge Header
            CopyMemory Buff(1), tBIH, Len(tBIH)
            'Get and merge DIB Bits
            GetDIBits lHDC, Picture, 0, tBI.bmHeight, Buff(Len(tBIH) + 1), tBIH, 0
            SelectObject lHDC, lhOldBmp
            DeleteDC lHDC
            'Create brush from packed DIB
            m_hBrush = CreateDIBPatternBrushPt(Buff(1), 0)
        End If
    End If
    '
    If (m_hBrush <> 0) Then
        SetRect TileRect, X1, Y1, X2, Y2
        SetBrushOrgEx hdc, X1, Y1, PtOrg
        'Tile image
        FillRect hdc, TileRect, m_hBrush
        DeleteObject m_hBrush
        m_hBrush = 0
    End If
    '
    i = SetStretchBltMode(hdc, lPicMode)
End Function
