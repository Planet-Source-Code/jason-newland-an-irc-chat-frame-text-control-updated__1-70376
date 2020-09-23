VERSION 5.00
Begin VB.UserControl ucChatFrame 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ToolboxBitmap   =   "ucChatFrame.ctx":0000
   Begin Project1.ucScrollbar vs 
      Height          =   2670
      Left            =   5040
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4710
      Enabled         =   0   'False
      Min             =   1
      Max             =   1
      Value           =   1
      Style           =   2
   End
   Begin VB.PictureBox pChat 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   30
      MouseIcon       =   "ucChatFrame.ctx":0312
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   332
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   4980
   End
End
Attribute VB_Name = "ucChatFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'IRC Chat Output Window Control (version 2.4 27 April, 2008)
'Original Concept: Vesa Piittinen <vesa@merri.net> - 2005-2007
'This Version:     Jason James Newland - 1st Feb, 2007-2008
'Other Credits:    Kim Tore Jensen - 2001-2003
'                  Marzo Junior, marzojr@taskmail.com.br, 20041027
'                  (Wordwrap routine)
'Requires FastString typelib (strings.tlb)
'
'Changes: (v.2.1) 19 April, 2008
'   - Paint timer now changed to stop updating the display when you are
'   selecting text to copy
'   - Changed Ctrl+C selection marking to now copy from the mouse start
'   position to the mouse end position (but taking into account any codes
'   in the line
'   - Removed the stupid vs_Scroll, Call vs_Scroll bottleneck, have no idea
'   why I put that, except it was to originally call vs_Change (this may
'   have caused a GPF on some systems when using the scrollbar
'Changes: (v.2.2) 21 April, 2008
'   - Fixed Ctrl+C bug not selecting start of mouse pos if only a single line
'   is selected (copied whole line)
'   - Made output unicode UTF-8 compatiable
'Changes: (v.2.3) 23 April, 2008
'   - Fixed line draw issue where colors and formatting would mess up when resizing the
'   the window
'   - Removed useless bits of code that were no longer needed (AddLine .Stripped property)
'   this will decrease the amount of memory usuage by 1/3
'   - Added back the vs_Scroll property "Call vs_Change"
'   - Fixed copy selection quirkiness on hanging (second or third) lines
'Changes: (v.2.4) 28 April, 2008
'   - Fixed a wrapping issue where it would overshoot the edge of the window if using
'   bolding and/or color formatting (now just minuses 32 of the total scalewidth)
'   - Re-Fixed the first fix from the 23 April change, on large amounts of text using
'   the vs_Scroll property or mousewheel would be hugely slow in drawing, so changed it
'   to search backwards a line to find the first line .FirstLine property and test draw
'   there on to the visible range
'   - Fixed a few other issues that I cant remember right now
'   - Added a "reload from log" feature which can load a predetermined (keep the number
'   low to increase speed) number of lines into the text buffer
'
Private Type LineData
    Text()                      As Byte     'Text data
    Color()                     As Byte     'Event color
End Type

Private Type WrapData
    Text()                      As Byte     'Text array for each line
    LN                          As Long     'Current line number text() points to in LineData
    FirstLine                   As Boolean  'Starting line of the wrap
    IsSet                       As Boolean  'Used for setting last used formatting
    ForeCol                     As Long     'Last known fore color
    BackCol                     As Long     'Last known back color
    bIsBold                     As Boolean  'Last known bolding
    bIsUnder                    As Boolean  'Last known underline
    bIsItalic                   As Boolean  'Last known italics
End Type

Private Type NickNames
    Nick()                      As Byte
End Type

'Mouse constants
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_RBUTTONUP      As Long = &H205

Private m_Font                  As Font
Private TextArea                As RECT     'Main DrawText rectangle
Private lTextHeight             As Long
Private lTopLine                As Long
Private lWidth                  As Long

Private lFindText               As Long

Private colFront                As Long
Private colBack                 As Long

Private blInverse               As Boolean
Private colRevFront             As Long
Private colRevBack              As Long

'Marking variables
Private IsMarking               As Boolean
Private sLineText               As String
Private MarkCurLine             As Long
Private MarkCopyColors          As Boolean
Private MarkStartLine           As Long
Private MarkStartPos            As Long
Private MarkEndLine             As Long
Private MarkEndPos              As Long
Private MarkEndLineOld          As Long
Private MarkEndPosOld           As Long

'Wordwrap headers
Private Header1(5) As Long
Private Header2(5) As Long
Private SafeArray1() As Integer
Private SafeArray2() As Integer

'Word detection (URL, etc.)
Private lOldLineLoc             As Long
Private sLink()                 As Byte
Private sChan()                 As Byte
Private sNick()                 As Byte
Private sNicks()                As NickNames

Private lLargeChange            As Long
Private iEventQueue             As Integer

'Line data
Private Lines()                 As LineData
Private Wrapped()               As WrapData

'Subclassing
Private m_emr                   As EMsgResponse
Private d_msg                   As New MGSubclass

'BG Picture pointer
Private pPic                    As PictureBox

'Timers
Private WithEvents tmrPaint     As CLiteTimer
Attribute tmrPaint.VB_VarHelpID = -1
Private WithEvents tmrRefresh   As CLiteTimer
Attribute tmrRefresh.VB_VarHelpID = -1
Private WithEvents tmrWrap      As CLiteTimer
Attribute tmrWrap.VB_VarHelpID = -1

'Exposed events of this class
Public Event Click()
Public Event RClick()
Public Event DblClickLink(sLink() As Byte)
Public Event DblClickChan(sChan() As Byte)
Public Event DblClickNick(sNick() As Byte)
Public Event DblClick()
Public Event LogOut(sData() As Byte)
Public Event CopyComplete()

'Wraptext API
Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)

'Subclassing structure
Implements MISubclass

'Subclassing method
Private Property Let MISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    'no choice, must be present
End Property

Private Property Get MISubclass_MsgResponse() As EMsgResponse
    MISubclass_MsgResponse = emrPostProcess
End Property

Private Function MISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    'The main window callback routine for receiving
    'the attached message events
    Dim i As Long, j As Long, L As Long, strLine() As Byte, sWord() As Byte, sNickTmp() As Byte
    Dim X As Single, Y As Single, sTmp As String, s As String, lPos As Long, lFore As Long, lBack As Long, strTemp() As String
    Dim lTmp As Long, sColTmp As String, bBold As Boolean, bUnder As Boolean, bItalic As Boolean, bReverse As Boolean
    '
    Select Case iMsg
        Case WM_LBUTTONDOWN
            If (Not Wrapped) = True Then GoTo ExitMarkSel
            '
            If wParam = 9 Then
                'ctrl key is down
                MarkCopyColors = True
            ElseIf wParam = 1 Then
                MarkCopyColors = False
            Else
                GoTo ExitMarkSel
            End If
            '
            X = LoWord(lParam)
            Y = HiWord(lParam)
            MarkEndPosOld = 0
            MarkEndLineOld = 0
            MarkStartLine = LineLoc(Y) - 1
            If (MarkStartLine < 0) Or (MarkStartLine > lLargeChange) Then GoTo ExitMarkSel
            MarkEndLine = MarkStartLine
            sLineText = CStr(Wrapped(LnNum(MarkStartLine)).Text)
            MarkCurLine = LnNum(MarkStartPos)
            MarkStartPos = CharPos(sLineText, CLng(X), MarkCurLine) - IIf(Wrapped(LnNum(MarkStartLine)).FirstLine = False, 1, 0)
            If MarkStartPos = 0 Then GoTo ExitMarkSel
            MarkEndPos = MarkStartPos + 1
            If MarkEndPos > Len(sLineText) Then MarkEndPos = Len(sLineText)
            IsMarking = True
            pChat.MousePointer = vbIbeam
ExitMarkSel:
            '
        Case WM_MOUSEMOVE
            X = LoWord(lParam)
            Y = HiWord(lParam)
            '
            If IsMarking Then
                s = vbNullString
                '
                MarkEndLine = LineLoc(Y) - 1
                If MarkEndLine < MarkStartLine Then GoTo ExitMark
                If (MarkEndLine < 0) Or (MarkEndLine > lLargeChange) Then GoTo ExitMark
                MarkEndPos = CharPos(sLineText, CLng(X), MarkCurLine) + 1
                '
                If MarkEndPos = MarkEndPosOld Then
                    If MarkEndLine = MarkEndLineOld Then GoTo ExitMark
                End If
                For L = MarkEndLine To MarkEndLineOld + 1
                    Draw LnNum(L), L + 1, True
                Next
                If MarkEndLine = MarkStartLine Then
                    'Single line only
                    i = LnNum(MarkStartLine)
                    MarkCurLine = i
                    If i < 0 Then GoTo ExitMark
                    sLineText = CStr(Wrapped(LnNum(MarkStartLine)).Text)
                    Draw i, MarkStartLine + 1, True, MarkStartPos, MarkEndPos
                Else
                    'Multiple lines
                    i = LnNum(MarkStartLine)
                    If i < 0 Then GoTo ExitMark
                    s = StripCC(CStr(Wrapped(i).Text), "CURBIN")
                    If (MarkEndPos < MarkEndPosOld) Or (L > 0) Then Draw LnNum(MarkEndLine), MarkEndLine + 1, True
                    Draw LnNum(MarkStartLine), MarkStartLine + 1, True, MarkStartPos, Len(s)
                    For L = MarkStartLine + 1 To MarkEndLine - 1
                        s = StripCC(CStr(Wrapped(LnNum(L)).Text), "CURBIN")
                        Draw LnNum(L), L + 1, True, 1, Len(s)
                    Next L
                    i = LnNum(MarkEndLine)
                    MarkCurLine = i
                    If i > UBound(Wrapped) Then GoTo ExitMark
                    sLineText = CStr(Wrapped(i).Text)
                    If MarkEndPos < 1 Then GoTo ExitMark
                    Draw i, MarkEndLine + 1, True, 1, MarkEndPos
                End If
                '
                MarkEndPosOld = MarkEndPos
                MarkEndLineOld = MarkEndLine
ExitMark:
                '
            Else
                If lOldLineLoc <> LnNum(LineLoc(Y)) - 1 Then
                    lOldLineLoc = LnNum(LineLoc(Y)) - 1
                    pChat.MousePointer = vbDefault
                    Erase sLink(), sChan(), sNick(), sWord()
                End If
                'Detection code
                If lOldLineLoc > -1 Then
                    'strLine = StripCC(CStr(Wrapped(lOldLineLoc).Text), "CURBIN")
                    strLine = CStr(Wrapped(lOldLineLoc).Text)
                    sWord = WordUnderMouse(StripCC(CStr(strLine), "CURBIN"), CharPos(CStr(strLine), CLng(X), lOldLineLoc))
                    If LenB(Trim$(CStr(sWord))) <> 0 Then
                        sTmp = IIf(InStr(CStr(sWord), CC_1), Replace$(CStr(sWord), CC_1, vbNullString), CStr(sWord))
                        sNickTmp = RemChars(sWord)
                        If Left$(sTmp, 1) = ChrW$(35) Then
                            'Check channel names
                            sChan = sTmp
                            pChat.MousePointer = vbCustom
                            GoTo SubExit
                        ElseIf GetNickPos(sNickTmp) <> -1 Then
                            'Check nicks
                            sNick = sNickTmp
                            pChat.MousePointer = vbCustom
                            Erase sNickTmp()
                            GoTo SubExit
                        Else
                            'Check URLS
                            If LCase$(sTmp) Like "*www.*.com*" Or LCase$(sTmp) Like "http:*.com*" Or LCase$(sTmp) Like "http:*" Or LCase$(sTmp) Like "*.net*" Or LCase$(sTmp) Like "*.org*" Then
                                GoTo url
                            Else
                                GoTo SubExit
                            End If
url:
                            'Somehow detect if the line was complete, if not
                            'add the next line to it to complete the url
                            L = 0
                            While Right$(s, 1) <> CC_1
                                s = s & CStr(Wrapped(lOldLineLoc + L).Text)
                                L = L + 1
                            Wend
                            s = StripCC(s, "CURBIN")
                            L = InStr(s, sTmp)
                            If L <> 0 Then
                                sTmp = GetTok(Mid$(s, L), "1", 32)
                            End If
                            sLink = IIf(InStr(sTmp, CC_1), Replace$(sTmp, CC_1, vbNullString), sTmp)
                            pChat.MousePointer = vbCustom
                        End If
                    Else
                        Erase sLink(), sChan(), sNick(), sWord()
                        pChat.MousePointer = vbDefault
                        GoTo SubExit
                    End If
                End If
SubExit:
            End If
        Case WM_LBUTTONUP
            If IsMarking = True Then
                pChat.MousePointer = vbDefault
                s = vbNullString
                '
                If MarkCopyColors Then
                    If MarkStartLine <= MarkEndLine Then
                        For i = MarkStartLine To MarkEndLine
                            L = LnNum(i)
                            If (L > UBound(Wrapped)) Or (L < 0) Then Exit For
                            sTmp = CStr(Wrapped(LnNum(i)).Text)
                            lPos = 0
                            sColTmp = vbNullString
                            If i = MarkStartLine Or i = MarkEndLine Then
                                For j = 1 To Len(sTmp)
                                    Select Case Mid$(sTmp, j, 1)
                                        Case CC_COLOR
                                            lTmp = j
                                            GetColors Mid$(sTmp, j + 1, 5), j, lFore, lBack
                                            If lPos < MarkStartPos And i = MarkStartLine Then
                                                sColTmp = Mid$(sTmp, lTmp, (j - lTmp) + 1)
                                            Else
                                                s = s & Mid$(sTmp, lTmp, (j - lTmp) + 1)
                                            End If
                                        Case CC_BOLD
                                            If lPos < MarkStartPos And i = MarkStartLine Then
                                                bBold = Not bBold
                                            Else
                                                s = s & Mid$(sTmp, j, 1)
                                            End If
                                        Case CC_UNDER
                                            If lPos < MarkStartPos And i = MarkStartLine Then
                                                bUnder = Not bUnder
                                            Else
                                                s = s & Mid$(sTmp, j, 1)
                                            End If
                                        Case CC_ITALIC
                                            If lPos < MarkStartPos And i = MarkStartLine Then
                                                bItalic = Not bItalic
                                            Else
                                                s = s & Mid$(sTmp, j, 1)
                                            End If
                                        Case CC_NORMAL
                                            If lPos >= MarkStartPos And i = MarkStartLine Then
                                                s = s & Mid$(sTmp, j, 1)
                                            End If
                                        Case CC_INVERSE
                                            If lPos < MarkStartPos And i = MarkStartLine Then
                                                bReverse = Not bReverse
                                            Else
                                                s = s & Mid$(sTmp, j, 1)
                                            End If
                                        Case Else
                                            lPos = lPos + 1
                                            If lPos >= MarkStartPos And i = MarkStartLine Then
                                                s = s & Mid$(sTmp, j, 1)
                                            ElseIf i = MarkEndLine And i <> MarkStartLine Then
                                                s = s & Mid$(sTmp, j, 1)
                                            End If
                                            If lPos >= MarkEndPos - 1 And i = MarkEndLine Then Exit For
                                    End Select
                                Next j
                                'Build the final string
                                s = IIf(LenB(sColTmp) <> 0, sColTmp, vbNullString) & IIf(bBold = True, CC_BOLD, vbNullString) & IIf(bUnder = True, CC_UNDER, vbNullString) & IIf(bItalic = True, CC_ITALIC, vbNullString) & IIf(bReverse = True, CC_INVERSE, vbNullString) & s
                            Else
                                s = s & sTmp
                            End If
                        Next
                    End If
                Else
                    If MarkStartLine >= MarkEndLine Then
                        If MarkEndPos > MarkStartPos Then 'Copy event
                            i = LnNum(MarkStartLine)
                            sTmp = Mid$(StripCC(CStr(Wrapped(i).Text), "CURBIN"), MarkStartPos, MarkEndPos - MarkStartPos)
                            s = sTmp
                        End If
                    Else
                        i = LnNum(MarkStartLine)
                        sTmp = StripCC(CStr(Wrapped(i).Text), "CURBIN")
                        s = Mid$(sTmp, MarkStartPos, Len(sTmp))
                        For i = MarkStartLine + 1 To MarkEndLine - 1
                            'Get all selected lines if more than one COMPLETE line
                            'is selected
                            If LnNum(i) > UBound(Wrapped) Then Exit For
                            sTmp = StripCC(CStr(Wrapped(LnNum(i)).Text), "CURBIN")
                            s = s & sTmp
                        Next
                        i = LnNum(MarkEndLine)
                        If i <= UBound(Wrapped) Then
                            'This appends the last part of the line and the end pos is going
                            'to always be one out (haven't figured out why ??) so, cheat and
                            'minus one
                            sTmp = Mid$(StripCC(CStr(Wrapped(LnNum(MarkEndLine)).Text), "CURBIN"), 1, MarkEndPos - 1)
                            s = s & sTmp
                        End If
                    End If
                End If
                'Seems like a long way around, but it makes sure the clipboard is first cleared
                'and sets the text correctly
                If LenB(s) <> 0 Then
                    strTemp = Split(s, CC_1)
                    If UBound(strTemp) > 0 Then
                        strLine = Replace$(s, CC_1, vbCrLf)
                    Else
                        strLine = s
                    End If
                    Clipboard.Clear
                    Clipboard.SetText StrConv(WToA(CStr(strLine), CP_UTF8), vbUnicode)
                    Erase strTemp(), strLine()
                End If
                '
                IsMarking = False
                MarkCurLine = 0
                MarkCopyColors = False
                MarkStartLine = 0
                MarkStartPos = 0
                MarkEndLine = 0
                MarkEndPos = 0
                MarkEndLineOld = 0
                MarkEndPosOld = 0
                Refresh
                RaiseEvent CopyComplete
            Else
                If LenB(CStr(sNick)) + LenB(CStr(sChan)) + LenB(CStr(sLink)) = 0 Then RaiseEvent Click
            End If
        Case WM_LBUTTONDBLCLK
            If LenB(CStr(sChan)) <> 0 Then
                RaiseEvent DblClickChan(sChan)
            ElseIf LenB(CStr(sNick)) <> 0 Then
                RaiseEvent DblClickNick(sNick)
            ElseIf LenB(CStr(sLink)) <> 0 Then
                RaiseEvent DblClickLink(sLink)
            Else
                RaiseEvent DblClick
            End If
        Case WM_RBUTTONUP
            RaiseEvent RClick
        Case WM_MOUSEWHEEL
            If vs.Enabled = True Then
                If Left$(Trim$(Str$(wParam)), 1) = "-" Then
                    i = vs.Value + 3
                Else
                    i = vs.Value - 3
                End If
                '
                If i <= vs.Max Then
                    vs.Value = i
                Else
                    vs.Value = vs.Max
                End If
            End If
    End Select
    '
    'I like kangaroos :)
End Function

'Private control methods
Private Sub pChat_Paint()
    On Error Resume Next
    Dim i As Long, j As Long, k As Long, t As Long
    '
    If (Not Wrapped) = True Then
        'just draw picture
        DrawBGPicture
        Exit Sub
    End If
    iEventQueue = 0
    j = vs.Value

    'draw picture (causes picturebox to clear)
    DrawBGPicture
    If j > lLargeChange Then
        k = -1
    ElseIf j <= lLargeChange Then
        k = lLargeChange - j
    ElseIf j = lLargeChange Then
        k = 0
    End If
    '
    For i = j - lLargeChange + k - 1 To 0 Step -1
        If Wrapped(i).FirstLine = True Then
            For t = i To j - lLargeChange + k - 1
                Draw t, -1
            Next t
            Exit For
        End If
    Next i
    '
    For i = k + 1 To lLargeChange
        Draw j - lLargeChange + i - 1, i
    Next
End Sub

Private Sub pChat_Resize()
    On Error Resume Next
    Dim i As Long
    '
    vs.Move UserControl.ScaleWidth - vs.Width, 0, vs.Width, UserControl.ScaleHeight
    '
    If lWidth <> pChat.ScaleWidth Then
        lWidth = pChat.ScaleWidth
        'Only wrap if the width of the window has changed, not the height (speed)
        Set tmrWrap = New CLiteTimer
        tmrWrap.Interval = 10
        tmrWrap.Enabled = True
        'Refresh
        Refresh
    Else
        'Just refresh
        Refresh
    End If
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    'Subclass
    d_msg.Attach_Message Me, pChat.hwnd, WM_LBUTTONDOWN
    d_msg.Attach_Message Me, pChat.hwnd, WM_LBUTTONDBLCLK
    d_msg.Attach_Message Me, pChat.hwnd, WM_MOUSEMOVE
    d_msg.Attach_Message Me, pChat.hwnd, WM_LBUTTONUP
    d_msg.Attach_Message Me, pChat.hwnd, WM_RBUTTONUP
    d_msg.Attach_Message Me, pChat.hwnd, WM_MOUSEWHEEL
    '
    Set Font = UserControl.Font
    '
    lOldLineLoc = -1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        Set m_Font = .ReadProperty("Font", UserControl.Font)
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    pChat.Move 2, 2, (ScaleWidth - 4) - vs.Width, ScaleHeight - 4
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    'Unsubclass!!
    d_msg.Detach_Message Me, pChat.hwnd, WM_LBUTTONDOWN
    d_msg.Detach_Message Me, pChat.hwnd, WM_LBUTTONDBLCLK
    d_msg.Detach_Message Me, pChat.hwnd, WM_MOUSEMOVE
    d_msg.Detach_Message Me, pChat.hwnd, WM_LBUTTONUP
    d_msg.Detach_Message Me, pChat.hwnd, WM_RBUTTONUP
    d_msg.Detach_Message Me, pChat.hwnd, WM_MOUSEWHEEL
    '
    Erase Lines(), Wrapped(), sLink(), sNick(), sChan(), sNicks()
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "Font", m_Font, UserControl.Font
    End With
End Sub

'Drawing routines
Private Sub Draw(lNum As Long, lPos As Long, Optional ClearRECT As Boolean, Optional lStart As Long = -1, Optional lEnd As Long = -1)
    On Error Resume Next
    Dim tmpLen As Long, sTemp As String, strTemp As String, i As Long, L As Long, t As Long, ccCodeBrush As Long, blBold As Boolean
    '
    'Set colors
    If Wrapped(lNum).FirstLine = True Then
        pChat.ForeColor = IRCColor(CInt(EventColor(Lines(Wrapped(lNum).LN).Color(0)) And &HF))
        pChat.FontBold = False
        pChat.FontUnderline = False
        pChat.FontItalic = False
        colFront = -1
        colBack = -1
        blInverse = False
    Else
        With Wrapped(lNum)
            If Not .IsSet Then
                .IsSet = True
                .bIsBold = pChat.FontBold
                .bIsItalic = pChat.FontItalic
                .bIsUnder = pChat.FontUnderline
                .ForeCol = colFront
                .BackCol = colBack
            Else
                pChat.FontBold = .bIsBold
                pChat.FontItalic = .bIsItalic
                pChat.FontUnderline = .bIsUnder
                colFront = .ForeCol
                colBack = .BackCol
                pChat.ForeColor = colFront
            End If
        End With
    End If
    'Resize drawing area (turn bolding off first so the left hanging justify appears
    '   the same no matter what the bold type)
    blBold = pChat.FontBold
    pChat.FontBold = False
    With TextArea
        .Left = IIf(Wrapped(lNum).FirstLine = True, 0, TextWidthU(Space$(2)))
        .Top = (lPos * lTextHeight) + lTopLine
        .Bottom = .Top + lTextHeight
        .Right = pChat.ScaleWidth
    End With
    'Reset bolding
    pChat.FontBold = blBold
    'Clear the line area if we are marking
    If lPos <> -1 Then
        If ClearRECT = True Then
            DeleteObject ccCodeBrush
            ccCodeBrush = CreateSolidBrush(pChat.BackColor)
            FillRect pChat.hDC, TextArea, ccCodeBrush
            DeleteObject ccCodeBrush
        End If
    End If
    'Prepare text output
    strTemp = RTrim$(CStr(Wrapped(lNum).Text))
    tmpLen = Len(strTemp)
    L = 0
    'If no formatting is in the line, skip processing (speed) [only if we're not marking]
    If Not IsMarking Then
        If CodeCount(strTemp) = 0 Then
            If InStr(strTemp, CC_1) Then
                strTemp = Replace$(strTemp, CC_1, vbNullString)
                tmpLen = Len(strTemp)
            End If
            If colFront <> -1 Then
                DoColors strTemp, colFront, colBack
            Else
                pChat.ForeColor = IRCColor(CInt(EventColor(Lines(Wrapped(lNum).LN).Color(0)) And &HF))
            End If
            DrawTextStr strTemp, TextArea, lPos
            GoTo SkipIT
        End If
    Else
        With Wrapped(lNum)
            If Not .FirstLine Then
                If .IsSet Then
                    pChat.FontBold = .bIsBold
                    pChat.FontItalic = .bIsItalic
                    pChat.FontUnderline = .bIsUnder
                    colFront = .ForeCol
                    colBack = .BackCol
                    pChat.ForeColor = colFront
                End If
            End If
        End With
    End If
    'Begin parsing line data character by character and draw the text
    For i = 1 To tmpLen
        'If no formatting is in the rest of the line, skip processing (speed) [only if we're not marking]
        sTemp = Mid$(strTemp, i)
        If Not IsMarking Then
            If CodeCount(sTemp) = 0 Then
                If InStr(sTemp, CC_1) Then sTemp = Replace$(sTemp, CC_1, vbNullString)
                If colFront <> -1 Then
                    DoColors sTemp, colFront, colBack
                Else
                    pChat.ForeColor = IRCColor(CInt(EventColor(Lines(Wrapped(lNum).LN).Color(0)) And &HF))
                End If
                '
                DrawTextStr sTemp, TextArea, lPos
                Exit For
            End If
        End If
        '
        Select Case Mid$(strTemp, i, 1)
            Case CC_COLOR
                'Color data
                If Mid$(strTemp, i + 1, 1) <> ChrW$(32) And LenB(Mid$(strTemp, i + 1, 1)) <> 0 Then
                    GetColors Mid$(strTemp, i + 1, 5), i, colFront, colBack
                    If colFront = -1 Then
                        'If ctrl+k character is on its own, works the same was as normalize
                        'revert back to default event color
                        colFront = IRCColor(CInt(EventColor(Lines(Wrapped(lNum).LN).Color(0)) And &HF))
                    Else
                        'Set the color (mainly the background rectangle)
                        DoColors Mid$(strTemp, i, 1), colFront, colBack
                    End If
                Else
                    'Its on its own
                    colFront = -1
                    colBack = -1
                    pChat.ForeColor = IRCColor(CInt(EventColor(Lines(Wrapped(lNum).LN).Color(0)) And &HF))
                End If
            Case CC_BOLD
                'Bold on/off
                pChat.FontBold = Not pChat.FontBold
            Case CC_UNDER
                'Underline on/of
                pChat.FontUnderline = Not pChat.FontUnderline
            Case CC_INVERSE
                'Inverse (reverse) on/off
                blInverse = Not blInverse
                If blInverse = True Then
                    colRevFront = colFront
                    colRevBack = colBack
                    colFront = IIf(colRevBack <> -1, colRevBack, pChat.BackColor)
                    colBack = IIf(colRevFront <> -1, colRevFront, pChat.ForeColor)
                Else
                    colFront = colRevFront
                    pChat.ForeColor = IIf(colFront <> -1, colFront, IRCColor(CInt(EventColor(Lines(Wrapped(lNum).LN).Color(0)) And &HF)))
                    colBack = colRevBack
                    colRevFront = -1
                    colRevBack = -1
                End If
            Case CC_ITALIC
                'Italics on/off (not supported by mIRC)
                pChat.FontItalic = Not pChat.FontItalic
            Case CC_NORMAL
                'Normal data (turn all formatting off)
                colFront = -1
                colBack = -1
                pChat.ForeColor = IRCColor(CInt(EventColor(Lines(Wrapped(lNum).LN).Color(0)) And &HF))
                If pChat.FontBold = True Then pChat.FontBold = False
                If pChat.FontUnderline = True Then pChat.FontUnderline = False
                If pChat.FontItalic = True Then pChat.FontItalic = False
            Case CC_1
                'End of line character (turn all formatting off)
                colFront = -1
                colBack = -1
                If pChat.FontBold = True Then pChat.FontBold = False
                If pChat.FontUnderline = True Then pChat.FontUnderline = False
                If pChat.FontItalic = True Then pChat.FontItalic = False
            Case Else
                'No special codes, draw text keeping in mind the color formatting
                If lPos <> -1 Then
                    L = L + 1
                    'If out lPos is -1 then we are only "test" drawing, or making sure that
                    'our picture window maintains the correct line colors and formatting
                    '23 April, 2008
                    If L >= lStart And L <= lEnd - 1 Then
                        With TextArea
                            t = .Right
                            .Right = .Left + TextWidthU(Mid$(strTemp, i, 1))
                            DeleteObject ccCodeBrush
                            ccCodeBrush = CreateSolidBrush(IRCColor(CInt(EventColor(13))))
                            FillRect pChat.hDC, TextArea, ccCodeBrush
                            DeleteObject ccCodeBrush
                            .Right = t
                        End With
                        pChat.ForeColor = pChat.BackColor
                    Else
                        If Not IsMarking Then
                            DoColors Mid$(strTemp, i, 1), colFront, colBack
                        Else
                            If colFront = -1 Then
                                pChat.ForeColor = IRCColor(CInt(EventColor(Lines(Wrapped(lNum).LN).Color(0)) And &HF))
                            Else
                                DoColors Mid$(strTemp, i, 1), colFront, colBack
                            End If
                        End If
                    End If
                    '
                    DrawTextStr Mid$(strTemp, i, 1), TextArea, lPos
                    TextArea.Left = TextArea.Left + TextWidthU(Mid$(strTemp, i, 1))
                End If
        End Select
    Next i
SkipIT:
    'refresh if required
    If lPos <> -1 Then
        If pChat.AutoRedraw Then
            Set tmrRefresh = New CLiteTimer
            tmrRefresh.Interval = 5
            tmrRefresh.Enabled = True
        End If
    End If
    '
    DeleteObject ccCodeBrush
End Sub

Private Sub DrawTextStr(ByVal sText As String, RC As RECT, lPos As Long)
    On Error Resume Next
    '
    If lPos <> -1 Then
        If Not b_IsNT Then
            '98 based, draw ANSI
            DrawText pChat.hDC, sText, -1, RC, DT_NOCLIP Or DT_NOPREFIX
        Else
            'NT based, draw unicode
            DrawTextUnicode pChat.hDC, StrPtr(sText), -1, RC, DT_NOCLIP Or DT_NOPREFIX
        End If
    End If
End Sub

Private Function TextWidthU(ByVal sString As String) As Long
    Dim TextRect As RECT
    SetRect TextRect, 0, 0, 0, 0
    If Not b_IsNT Then
        DrawText UserControl.hDC, sString, -1, TextRect, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_LEFT
    Else
        DrawTextUnicode UserControl.hDC, StrPtr(sString), -1, TextRect, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_LEFT
    End If
    TextWidthU = TextRect.Right
End Function

Private Function CodeCount(sText As String) As Long
    On Error Resume Next
    CodeCount = InStr(sText, CC_COLOR) + InStr(sText, CC_BOLD) + InStr(sText, CC_UNDER) + InStr(sText, CC_ITALIC) + InStr(sText, CC_INVERSE) + InStr(sText, CC_NORMAL)
End Function

Private Sub DoColors(ByVal s As String, ForeColor As Long, BackColor As Long)
    On Error Resume Next
    Dim tRight As Long, tBottom As Long, ccCodeBrush As Long
    If ForeColor <> -1 Then pChat.ForeColor = ForeColor
    '
    If BackColor <> -1 Then
        'If the back color is present, create a drawing brush
        ccCodeBrush = CreateSolidBrush(BackColor)
        With TextArea
            'Store original rectangle dimensions
            tRight = .Right
            tBottom = .Bottom
            'Modify them
            .Right = .Left + TextWidthU(s)
            .Bottom = .Top + pChat.TextHeight(s)
            'Fill the background area with the back color
            FillRect pChat.hDC, TextArea, ccCodeBrush
            'DELETE THE OBJECT AFTER ITS FINISHED WITH
            DeleteObject ccCodeBrush
            'Reset the dimensions of the rectangle
            .Right = tRight
            .Bottom = tBottom
        End With
    End If
End Sub

Private Sub GetColors(ByVal s As String, ByRef NextPos As Long, ByRef ForeColor As Long, ByRef BackColor As Long, Optional iLastFore As Integer = -1)
    On Error Resume Next
    Dim M As String, C(0 To 1) As Long
    '
    ForeColor = -1
    C(0) = -1
    C(1) = -1
    M = Left$(s, 2)
    If (Right$(M, 1) = "-") Or (Right$(M, 1) = "+") Or (Right$(M, 1) = ",") Or (Right$(M, 1) = " ") Then M = Left$(s, 1)
    If Not IsNumeric(M) Then M = Left$(s, 1)
    If IsNumeric(M) Then
        'Foreground color given
        C(0) = CLng(M)
        NextPos = NextPos + Len(M)
        s = Mid$(s, Len(M) + 1)
        M = Left$(s, 1)
        If M = "," Then
            'Background color possibly given
            M = Mid$(s, 2, 2)
            If (Right$(M, 1) = "-") Or (Right$(M, 1) = "+") Or (Right$(M, 1) = ",") Or (Right$(M, 1) = " ") Then M = Left$(M, 1)
            If Not IsNumeric(M) Then M = Left$(M, 1)
            If IsNumeric(M) Then
                'Background color IS given
                C(1) = CLng(M)
                NextPos = NextPos + Len(M) + 1
            End If
        End If
    Else
        'No background color
        BackColor = -1
        Exit Sub
    End If
    'Return the color information as a long
    If (C(0) > -1) Then If (C(0) < 16) Then ForeColor = IRCColor(C(0))
    If (C(1) > -1) Then If (C(1) < 16) Then BackColor = IRCColor(C(1))
    If (C(0) > 15) Then If (C(0) < 32) Then ForeColor = IRCColor(C(0) - 16)
    If (C(1) > 15) Then If (C(1) < 32) Then BackColor = IRCColor(C(1) - 16)
End Sub

'Text wrapping
Private Sub BuildWrap()
    On Error Resume Next
    Dim i As Long, bMax As Boolean, sText As String, t As Integer, L As Long, strTemp() As String
    Dim tmpArr() As WrapData, lTotal As Long
    Static HereAlready As Boolean
    '
    If HereAlready = True Then Exit Sub
    pChat.MousePointer = vbHourglass
    lTotal = UBound(Lines)
    i = 0
    HereAlready = True
    '
    Do
        If lTotal < UBound(Lines) Then lTotal = UBound(Lines)
        DoEvents
        If i > lTotal Then
            'Erase original array
            Erase Wrapped()
            'Set wrapped array to the temp array
            Wrapped = tmpArr
            Erase tmpArr()
            'Re-adjust the scrollbar
            If vs.Value = vs.Max Then bMax = True
            vs.Min = 1
            vs.Max = UBound(Wrapped) + 1
            If bMax = True Then vs.Value = vs.Max
            'Refresh display
            Refresh
            pChat.MousePointer = vbDefault
            HereAlready = False
            Exit Do
        End If
        '
        If (Not tmpArr) = True Then
            L = -1
        Else
            L = UBound(tmpArr)
        End If
        '
        sText = Trim$(CStr(Lines(i).Text))
        strTemp = Split(WrapText(sText & CC_1), vbCrLf)
        '
        For t = LBound(strTemp) To UBound(strTemp)
            If Not IsNull(strTemp(t)) Then
                ReDim Preserve tmpArr(L + 1 + t)
                If t = 0 Then
                    tmpArr(L + 1 + t).FirstLine = True
                    tmpArr(L + 1 + t).IsSet = False
                Else
                    tmpArr(L + 1 + t).FirstLine = False
                    tmpArr(L + 1 + t).IsSet = False
                End If
                tmpArr(L + 1 + t).Text = strTemp(t)
                tmpArr(L + 1 + t).LN = i
            Else
                'If it was the last line, we need to put the end of line character
                'on the end of the previous line
                If t = UBound(strTemp) Then
                    If Right$(CStr(tmpArr(L + t).Text), 1) <> CC_1 Then tmpArr(L + t).Text = CStr(tmpArr(L + t).Text) & CC_1
                End If
            End If
        Next t
        Erase strTemp()
        i = i + 1
    Loop
End Sub

Private Sub Wrap(ByVal sText As String, lNum As Long)
    On Error Resume Next
    Dim i As Long, t As Integer, L As Long, strTemp() As String
    If (Not Wrapped) = True Then
        L = -1
    Else
        L = UBound(Wrapped)
    End If
    '
    strTemp = Split(WrapText(sText & CC_1), vbCrLf)
    '
    For t = LBound(strTemp) To UBound(strTemp)
        If Not IsNull(strTemp(t)) Then
            ReDim Preserve Wrapped(L + 1 + t)
            If t = 0 Then
                Wrapped(L + 1 + t).FirstLine = True
                Wrapped(L + 1 + t).IsSet = False
            Else
                Wrapped(L + 1 + t).FirstLine = False
                Wrapped(L + 1 + t).IsSet = False
            End If
            Wrapped(L + 1 + t).Text = strTemp(t)
            Wrapped(L + 1 + t).LN = lNum
        Else
            'If it was the last line, we need to put the end of line character
            'on the end of the previous line
            If t = UBound(strTemp) Then
                If Right$(CStr(Wrapped(L + t).Text), 1) <> CC_1 Then Wrapped(L + t).Text = CStr(Wrapped(L + t).Text) & CC_1
            End If
        End If
    Next t
    Erase strTemp()
End Sub

Private Function IsNull(ByVal sText As String) As Boolean
    On Error Resume Next
    Dim strTemp As String
    '
    strTemp = sText
    If InStr(strTemp, CC_BOLD) Then strTemp = Replace$(strTemp, CC_BOLD, vbNullString)
    If InStr(strTemp, CC_NORMAL) Then strTemp = Replace$(strTemp, CC_COLOR, vbNullString)
    If InStr(strTemp, CC_UNDER) Then strTemp = Replace$(strTemp, CC_UNDER, vbNullString)
    If InStr(strTemp, CC_INVERSE) Then strTemp = Replace$(strTemp, CC_INVERSE, vbNullString)
    If InStr(strTemp, CC_NORMAL) Then strTemp = Replace$(strTemp, CC_NORMAL, vbNullString)
    If InStr(strTemp, CC_ITALIC) Then strTemp = Replace$(strTemp, CC_ITALIC, vbNullString)
    If InStr(strTemp, CC_15) Then strTemp = Replace$(strTemp, CC_15, vbNullString)
    If InStr(strTemp, CC_1) Then strTemp = Replace$(strTemp, CC_1, vbNullString)
    '
    IsNull = IIf(LenB(StripCC(strTemp, "C")) = 0, True, False)
End Function

'Line and word detection functions
Private Function LnNum(ByVal LnPos As Long) As Long
    On Error Resume Next
    'LnNum = (vs.Value - lLargeChange) + LnPos
    LnNum = (vs.Value - lLargeChange) + LnPos
End Function

Private Function LineLoc(ByVal Y As Long) As Long
    On Error Resume Next
    LineLoc = (Y - lTopLine) \ lTextHeight
End Function

Private Function WordUnderMouse(ByVal sLine As String, lCharPos As Long) As String
    On Error Resume Next
    Dim pos As Long, i As Long, lStart As Long, lEnd As Long, lLen As Long, sTxt As String, sChr As String
    '
    lLen = Len(sLine)
    sTxt = sLine
    pos = lCharPos
    If pos > 0 Then
        For i = pos To 1 Step -1
            sChr = Mid$(sTxt, i, 1)
            If sChr = " " Or sChr = vbCr Or i = 1 Then
                'If the starting character is vbCrLf then
                'we want to chop that off
                If sChr = vbCr Then
                    lStart = (i + 2)
                Else
                    lStart = i
                End If
                Exit For
            End If
        Next i
        '
        For i = pos To lLen
            If Mid$(sTxt, i, 1) = " " Or Mid$(sTxt, i, 1) = vbCr Or i = lLen Then
                lEnd = i + 1
                Exit For
            End If
        Next i
        '
        If lEnd >= lStart Then
            WordUnderMouse = Trim$(Mid$(sTxt, lStart, (lEnd - lStart)))
        End If
    End If
End Function

Private Function CharPos(ByVal sText As String, X As Long, lNum As Long) As Long
    On Error Resume Next
    'This sub will determine by character width what position in the string the
    'mouse pointer is over
    Dim i As Long, lLen As Long, t As Long, strTemp As String, lColPos As Long, L As Long, sTmp As String
    '
    L = 0
    lLen = 0
    sTmp = RTrim$(sText)
    '
    With Wrapped(lNum)
        If .FirstLine Then
            UserControl.FontBold = False
            UserControl.FontItalic = False
        Else
            L = TextWidthU(Space$(2))
            UserControl.FontBold = .bIsBold
            UserControl.FontItalic = .bIsItalic
        End If
    End With
    '
    For i = 1 To Len(sTmp)
        If L >= X Then
            CharPos = lLen
            Exit Function
        ElseIf i = Len(sTmp) And IsMarking = True Then
            CharPos = Len(sTmp)
            Exit Function
        Else
            Select Case Mid$(sTmp, i, 1)
                Case CC_COLOR
                    If Mid$(sTmp, i + 1, 1) <> ChrW$(32) And LenB(Mid$(sTmp, i + 1, 1)) <> 0 Then
                        strTemp = Mid$(sTmp, i + 1, 5)
                        'Now we should have the color string
                        t = 1
                        GetColors strTemp, t, -1, -1
                        i = i + (t - 1)
                    End If
                Case CC_BOLD
                    UserControl.FontBold = Not UserControl.FontBold
                Case CC_ITALIC
                    UserControl.FontItalic = Not UserControl.FontItalic
                'Case CC_UNDER
                '    UserControl.FontUnderline = Not UserControl.FontUnderline
                Case CC_UNDER, CC_NORMAL, CC_INVERSE, CC_15
                    'Do nothing
                Case Else
                    L = L + TextWidthU(Mid$(sTmp, i, 1))
                    lLen = lLen + 1
            End Select
        End If
    Next i
End Function

Private Function RemChars(sData() As Byte) As String
    On Error Resume Next
    Dim sTmp As String
    sTmp = CStr(sData)
    If InStr(sTmp, CC_1) <> 0 Then sTmp = Replace$(sTmp, CC_1, vbNullString)
    If InStr(sTmp, ChrW$(40)) <> 0 Then sTmp = Replace$(sTmp, ChrW$(40), vbNullString)
    If InStr(sTmp, ChrW$(41)) <> 0 Then sTmp = Replace$(sTmp, ChrW$(41), vbNullString)
    If InStr(sTmp, "<") <> 0 Then sTmp = Replace$(sTmp, "<", vbNullString)
    If InStr(sTmp, ">") <> 0 Then sTmp = Replace$(sTmp, ">", vbNullString)
    If InStr(sTmp, ":") <> 0 Then sTmp = Replace$(sTmp, ":", vbNullString)
    RemChars = sTmp
End Function

'Public exposed methods
Public Sub AddLine(ByRef Text() As Byte, Optional ByVal DefaultColor As Byte = &H1, Optional TimeStamp As Boolean = True)
    On Error Resume Next
    Dim index As Long, a As Long, strTemp As String, sLines() As String, newLN As Long
    Dim i As Long, tmpArr() As Byte, bFirst As Boolean, j As Integer, bOnce As Boolean, bMax As Boolean
    'Make sure we are getting something
    If (Not Text) = True Then Exit Sub
    '
    strTemp = Trim$(CStr(Text))
    sLines = Split(strTemp, vbCrLf)
    If UBound(sLines) > 0 Then
        For i = LBound(sLines) To UBound(sLines)
            tmpArr = sLines(i)
            AddLine tmpArr, DefaultColor, IIf(sLines(i) = LineSep, False, TimeStamp)
        Next i
        Erase tmpArr()
        Exit Sub
    End If
    '
    If (Not Lines) = True Then
        index = 0
    Else
        'Check that the line before this one isn't a line separator
        If Replace$(CStr(Lines(UBound(Lines)).Text), ChrW$(1), vbNullString) = LineSep And CStr(Text) = LineSep Then Exit Sub
        'Create a new line
        index = UBound(Lines) + 1
    End If
    'Increase line buffer (if greater than the max, trim off first line)
    If index > 1000 Then
        bOnce = False
        For a = LBound(Lines) + 1 To UBound(Lines)
            Lines(a - 1) = Lines(a)
            If bOnce = False And a - 1 = 0 Then
                'This is about the fastest way I know of to remove the wrapped version
                'of the line being removed as re-wrapping 1000+ lines every new line
                'slows it down a HECK of a lot
                j = 0
                bOnce = True
                For i = LBound(Wrapped) To UBound(Wrapped)
                    If Right(Wrapped(i).Text, 1) = CC_1 Then
                        j = j + 1
                        Exit For
                    Else
                        j = j + 1
                    End If
                Next i
                '
                newLN = 0
                For i = j To UBound(Wrapped)
                    Wrapped(i - j) = Wrapped(i)
                    Wrapped(i - j).LN = newLN
                    If Right$(CStr(Wrapped(i - j).Text), 1) = CC_1 Then newLN = newLN + 1
                Next i
                ReDim Preserve Wrapped(UBound(Wrapped) - j)
            End If
        Next a
        index = UBound(Lines)
        '
        If vs.Value = vs.Max Then bMax = True
        vs.Min = 1
        vs.Max = UBound(Wrapped) + 1
        If bMax = True Then vs.Value = vs.Max
    Else
        ReDim Preserve Lines(index)
    End If
    '
    With Lines(index)
        'Get a timestamp (if the argument was true)
        tmpArr = IIf(TimeStamp = True, GetTimeStamp, vbNullString)
        'Reserve memory for our needs
        ReDim .Text(UBound(tmpArr) + UBound(Text))
        ReDim .Color(0)
        .Text = CByte(tmpArr) & CByte(Text)
        Erase tmpArr()
        .Color(0) = DefaultColor
        Wrap Trim$(CStr(.Text)), index
        'raise log out event for logging
        tmpArr = .Text
        RaiseEvent LogOut(tmpArr)
        Erase tmpArr()
    End With
    'Keep the correct position
    If vs.Value = vs.Max Then
        'Increase max, stay at bottom (not scrolling)
        vs.Max = UBound(Wrapped) + 1
        vs.Value = vs.Max
        lFindText = vs.Value - 1
        Call vs_Change
    ElseIf vs.Max > 0 Then
        'We are scrolled above the bottom line
        Beep
        vs.Max = UBound(Wrapped) + 1
        Call vs_Change
    Else
        'Trigger Change
        Call vs_Change
    End If
    'Event Queue
    iEventQueue = iEventQueue + 1
    If iEventQueue > 20 Then
        iEventQueue = 0
        If Not IsMarking Then Refresh
    End If
End Sub

Public Sub LoadFile(ByVal sFile As String, lLines As Long)
    'Purpose of this sub is to load part of a log file to the window
    On Error GoTo continue
    Dim FNum As Integer, strBuff As String, strLines() As String, strTemp As String
    Dim i As Long, tmpArr() As Byte, lCol As Long, lPos As Long
    '
    FNum = FreeFile
    Open sFile For Binary Access Read Shared As #FNum
        strBuff = Space$(LOF(FNum))
        Get #FNum, , strBuff
        strLines = Split(strBuff, vbCrLf)
continue:
    Close #FNum
    '
    lPos = UBound(strLines) - lLines
    If lPos < 0 Or lPos = UBound(strLines) Then lPos = LBound(strLines)
    For i = lPos To UBound(strLines)
        'Now we add our strings
        strTemp = WToA(GetTok(strLines(i), "2-", 44), CP_ACP)
        tmpArr = AToW(strTemp, CP_UTF8)
        lCol = Val(GetTok(strLines(i), "1", 44))
        AddLine tmpArr, IIf(lCol = 0, 13, lCol), False
    Next i
    'We insert our line separator
    tmpArr = LineSep
    AddLine tmpArr, 13, False
    '
    Erase tmpArr(), strLines()
End Sub

Public Sub Clear()
    On Error Resume Next
    Erase Lines(), Wrapped()
    '
    vs.Min = 1
    vs.Max = 1
    vs.Value = 1
    vs.Enabled = False
    '
    pChat_Paint
End Sub

Public Sub FindText(ByVal sText As String, Optional iDirection As Integer = 0)
    '0 is up, 1 is down
    On Error Resume Next
    Dim i As Long, strTemp As String
    '
    If LenB(sText) = 0 Or vs.Enabled = False Then
        Beep
        Exit Sub
    End If
    '
    Select Case iDirection
        Case 0
            'Up
            lFindText = lFindText - 1
            If lFindText < 0 Then
                lFindText = 0
                Beep
            End If
            '
            For i = lFindText To LBound(Wrapped) Step -1
                strTemp = LCase$(StripCC(CStr(Wrapped(i).Text), "CURBIN"))
                If InStr(strTemp, LCase$(sText)) <> 0 Then
                    lFindText = i
                    vs.Value = i + 1
                    Exit Sub
                End If
            Next i
        Case 1
            'Down
            lFindText = lFindText + 1
            If lFindText >= UBound(Wrapped) Then
                lFindText = vs.Value - 1
                Beep
            End If
            '
            For i = lFindText To UBound(Wrapped)
                strTemp = LCase$(StripCC(CStr(Wrapped(i).Text), "CURBIN"))
                If InStr(strTemp, LCase$(sText)) <> 0 Then
                    lFindText = i
                    vs.Value = i + 1
                    Exit Sub
                End If
            Next i
    End Select
    '
    'Nothing found
    Beep
    lFindText = vs.Max - 1
End Sub

Public Property Set Font(ByVal NewFont As StdFont)
    On Error Resume Next
    '
    Set m_Font = NewFont
    Set pChat.Font = NewFont
    Set UserControl.Font = NewFont
    '
    lWidth = 0
    pChat_Resize
    '
    PropertyChanged "Font"
End Property

Public Property Get Font() As Font
    On Error Resume Next
    Set Font = m_Font
End Property

Public Sub Refresh()
    On Error Resume Next
    lTextHeight = pChat.TextHeight("W")
    lLargeChange = pChat.ScaleHeight \ lTextHeight
    lTopLine = (pChat.ScaleHeight Mod lTextHeight) - lTextHeight
    vs.LargeChange = lLargeChange
    pChat_Paint
End Sub

'BG Picture
Public Property Set BGSource(pSource As PictureBox)
    On Error Resume Next
    Set pPic = pSource
    '
    Refresh
End Property

Public Property Let BackColor(lColor As Long)
    On Error Resume Next
    pChat.BackColor = lColor
    UserControl.BackColor = lColor
    pChat_Paint
End Property

Private Sub DrawBGPicture()
    On Error Resume Next
    Dim a As Long, b As Long, C As Long, d As Long, i As Long, lPicMode As Long
    Dim lwWidth As Long, lHeight As Long, Left As Long, Top As Long
    'here is how we set the text ontop of the background image, unfortunately
    'this method requires that each time the control is painted that this sub
    'is called in order to repaint the background, may seem like the slow
    'approach, but as there is no way to make a picturebox transparent and the
    'image control doesn't have an hDC, we cant draw text directly on to the control
    'we read the "tag" and check which operation we need to do
    With pChat
        .Cls
        '
        lPicMode = SetStretchBltMode(.hDC, STRETCH_HALFTONE)
        pPic.BackColor = .BackColor
        '
        a = .ScaleWidth
        b = .ScaleHeight
        C = pPic.ScaleWidth
        d = pPic.ScaleHeight
        '
        Select Case CLng(pPic.Tag)
            Case 1
                'centered
                If .ScaleWidth > pPic.ScaleWidth Then Left = (.ScaleWidth \ 2 - pPic.ScaleWidth \ 2)
                If .ScaleHeight > pPic.ScaleHeight Then Top = (.ScaleHeight \ 2) - (pPic.ScaleHeight \ 2)
                BitBlt .hDC, Left, Top, pPic.ScaleWidth, pPic.ScaleHeight, pPic.hDC, 0, 0, vbSrcCopy
                i = SetStretchBltMode(.hDC, lPicMode)
            Case 2
                'stretchmode
                StretchBlt .hDC, 0, 0, a, b, pPic.hDC, 0, 0, C, d, vbSrcCopy
                i = SetStretchBltMode(.hDC, lPicMode)
            Case 3
                'tiled
                TileBitmap pPic.Picture, .hDC, 0, 0, a, b
            Case 4
                'photo
                If d > 135 Then
                    lwWidth = C / 3
                    lHeight = d / 3
                Else
                    lwWidth = C
                    lHeight = d
                End If
                '
                StretchBlt .hDC, a - lwWidth, 0, lwWidth, lHeight, pPic.hDC, 0, 0, C, d, vbSrcCopy
                i = SetStretchBltMode(.hDC, lPicMode)
        End Select
    End With
End Sub

'Scrollbar routines
Private Sub vs_Change()
    On Error Resume Next
    'Adds a small delay to the paint routine (especially if we quickly grab the scroller
    'and shift it suddenly upwards
    'If vs_Value = 0 Then vs_Value = 1
    '
    Set tmrPaint = New CLiteTimer
    tmrPaint.Interval = 10
    tmrPaint.Enabled = True
End Sub

Private Sub vs_Scroll()
    On Error Resume Next
    'You can choose whether to call vs_Change or not
    Call vs_Change
End Sub

'Wraptext routines
Private Sub InitArr()
    On Error Resume Next
    'Set up our template for looking at strings
    Header1(0) = 1              'Number of dimensions
    Header1(1) = 2              'Bytes per element (long = 4)
    Header1(4) = &H7FFFFFFF     'Array size
    'Force SafeArray1 to use Header1 as its own header
    RtlMoveMemory ByVal ArrPtr(SafeArray1), VarPtr(Header1(0)), 4
    'Set up our template for look at search text
    Header2(0) = 1                 'Number of dimensions
    Header2(1) = 2                 'Bytes per element (long = 4)
    Header2(4) = &H7FFFFFFF        'Array size
    'Force SafeArray1 to use Header1 as its own header
    RtlMoveMemory ByVal ArrPtr(SafeArray2), VarPtr(Header2(0)), 4
End Sub

Private Sub DestroyArr()
    On Error Resume Next
    'Make SafeArray1 once again use its own header
    'If this code doesn't run the IDE will crash
    RtlMoveMemory ByVal ArrPtr(SafeArray1), 0&, 4
    RtlMoveMemory ByVal ArrPtr(SafeArray2), 0&, 4
    '
    Erase SafeArray1(), SafeArray2()
End Sub

Private Function WrapText(Text As String) As String
    On Error Resume Next
    'By Marzo Junior, marzojr@taskmail.com.br, 20041027
    'Based on the code by Donald Lessau
    'Note: this routine is not exactly fast (can take upto 3 seconds on 1000 lines
    ' to complete wrap - there isn't any other faster way). Keep this in mind when
    ' using the host form maximized, Windows, when changing form focus, sets every other
    ' form behind the active one to NORMAL window state, so the resize event is fired
    ' on each form.
    Dim i As Long, posBreak As Long, sTemp As String, strTemp As String, L As Long, lwWidth As Long
    Dim cntBreakChars As Long, ubText As Long, lLen As Long, t As Long, lColPos As Long, lIndent As Long
    '
    lwWidth = pChat.ScaleWidth - 32
    lIndent = 0
    UserControl.FontBold = False
    'Test length first to see if we need to wrap the line (speed)
    If CodeCount(Text) = 0 Then
        'No codes in line, test the rectangle
        L = TextWidthU(IIf(InStr(Text, CC_1), Replace$(Text, CC_1, vbNullString), Text))
        If L <= lwWidth - lIndent Then
            'No need to wrap
            WrapText = Text
            Exit Function
        End If
    End If
    'Initialize array
    InitArr
    '
    lLen = Len(Text)
    ubText = lLen - 1
    L = 0
    lColPos = 0
    'Point the arrays to our strings; also, allocate the potential max string:
    Header1(3) = StrPtr(Text)
    sTemp = FastString.SysAllocStringLen(ByVal 0, lLen * 3)
    Header2(3) = StrPtr(sTemp)
    For i = 0 To ubText
        'We could get even faster by testing the rest of the string for codes
        'and length
        strTemp = Mid$(Text, i + 1)
        If InStr(strTemp, CC_1) Then strTemp = Replace$(strTemp, CC_1, vbNullString)
        '
        If CodeCount(strTemp) = 0 And UserControl.FontBold = False Then
            t = L
            L = L + TextWidthU(strTemp)
            If L < lwWidth - lIndent - TextWidthU(Space$(1)) Then
                'no need to wrap any further
                WrapText = Left$(sTemp, i + cntBreakChars) & Mid$(Text, i + 1)
                DestroyArr
                Exit Function
            Else
                L = t
            End If
        End If
        '
        Select Case SafeArray1(i)
            Case 32, 33, 45 'Space, exclamation, hyphen
                posBreak = i
                L = L + TextWidthU(Mid$(Text, i + 1, 1))
            Case SP_BOLD
                'Bolding now seems to be shorter line length than the window
                'I can't figure out why or get it any closer than that
                UserControl.FontBold = Not UserControl.FontBold
            Case SP_UNDER, SP_ITALIC, SP_INVERSE, SP_NORMAL, 15, 1
                'Do nothing
            Case SP_COLOR
                'Here, the tricky bit is, we have to remove valid codes
                'from the line so we can correctly draw the line
                strTemp = Mid$(Text, i + 2, 5)
                t = 1
                GetColors strTemp, t, -1, -1
                lColPos = i + t
            Case Else
                If i >= lColPos Then L = L + TextWidthU(Mid$(Text, i + 1, 1))
        End Select
        '
        SafeArray2(i + cntBreakChars) = SafeArray1(i)
        If L > lwWidth - lIndent - TextWidthU(Space$(1)) Then
            If posBreak > 0 Then
                If posBreak > i - (i / 3) Then
                    'Don't break at the very end
                    If posBreak = ubText Then Exit For
                    'Wrap after space, hyphen
                    SafeArray2(posBreak + cntBreakChars + 1) = &HD
                    SafeArray2(posBreak + cntBreakChars + 2) = &HA
                    i = posBreak
                    posBreak = 0
                Else
                    'Cut word
                    SafeArray2(i + cntBreakChars) = &HD
                    SafeArray2(i + cntBreakChars + 1) = &HA
                    i = i - 1
                End If
            Else
                'Cut word
                SafeArray2(i + cntBreakChars) = &HD
                SafeArray2(i + cntBreakChars + 1) = &HA
                i = i - 1
            End If
            cntBreakChars = cntBreakChars + 2
            If lIndent = 0 Then
                lIndent = TextWidthU(Space$(2))
            End If
            L = 0
        End If
    Next i
    '
    WrapText = Left$(sTemp, lLen + cntBreakChars)
    DestroyArr
End Function

'Nickname storage functions
Public Sub AddNick(ByVal sNick As String)
    On Error Resume Next
    Dim lIndex As Long
    'first check the nick doesn't exist
    'redimension
    If (Not sNicks) = True Then
        lIndex = 0
    Else
        If GetNickPos(sNick) <> -1 Then Exit Sub
        lIndex = UBound(sNicks) + 1
    End If
    ReDim Preserve sNicks(lIndex)
    'add the nick name
    sNicks(lIndex).Nick = sNick
End Sub

Public Sub RemoveNick(ByVal sNick As String)
    On Error Resume Next
    Dim lIndex As Long, i As Long
    lIndex = GetNickPos(sNick)
    For i = lIndex + 1 To UBound(sNicks)
        sNicks(i - 1) = sNicks(i)
    Next i
    'redimension
    ReDim Preserve sNicks(UBound(sNicks) - 1)
End Sub

Public Sub ClearNicks()
    On Error Resume Next
    Erase sNicks()
End Sub

Private Function GetNickPos(ByVal sNick As String) As Long
    On Error Resume Next
    Dim i As Long
    If (Not sNicks) = True Then
        GetNickPos = -1
        Exit Function
    End If
    '
    For i = LBound(sNicks) To UBound(sNicks)
        If StrComp(CStr(sNicks(i).Nick), sNick, 1) = 0 Then
            GetNickPos = i
            Exit Function
        End If
    Next i
    GetNickPos = -1
End Function

'Timers
Private Sub tmrPaint_Timer()
    On Error Resume Next
    If Not IsMarking Then
        tmrPaint.Enabled = False
        '
        iEventQueue = 0
        Refresh
        '
        Set tmrPaint = Nothing
    End If
End Sub

Private Sub tmrRefresh_Timer()
    On Error Resume Next
    tmrRefresh.Enabled = False
    '
    If vs.Max >= lLargeChange Then
        If vs.Enabled = False Then vs.Enabled = True
    Else
        If vs.Enabled = True Then vs.Enabled = False
    End If
    '
    pChat.Refresh
    Set tmrRefresh = Nothing
End Sub

Private Sub tmrWrap_Timer()
    On Error Resume Next
    tmrWrap.Enabled = False
    '
    BuildWrap
    '
    Set tmrWrap = Nothing
End Sub
