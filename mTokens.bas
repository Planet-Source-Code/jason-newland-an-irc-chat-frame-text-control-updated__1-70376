Attribute VB_Name = "mTokens"
'Gettok functions, etc, for use with IRC scripting but
'may serve other purposes
Option Explicit

'GetTok("text with delimiters", "position as a string eg: 1", character (ie: 32))
Public Function GetTok(ByVal strSource As String, ByVal strPosition As String, ByVal Token As Integer, Optional ByVal NumTokens As Integer) As String
    'by Jason James Newland (2006)
    'eg: gettok("hi world this is a test","2-",32) would
    'return "world this is a test" and
    'gettok("hi world this is a test","-2",32) would
    'return "hi world"
    'gettok("hi world this is a test","0",32) would return
    '6 as in 6 tokens delimited by chr(32) in the source
    'string
    'gettok("hi world this is a test","2 to 4",32) would
    'return "world this is"
    On Error Resume Next
    Dim Tokens() As String, intPosition As Integer, intTemp As Integer, intPos As Integer
    Dim intBool As Boolean, intRev As Boolean, i As Long, strToken As String
    Dim NumOfToks As Integer, lLower As Long, lUpper As Long
    '
    'first if the position is 0 return the total number of tokens
    If CLng(Replace$(strPosition, "-", vbNullString)) = 0 Then
        Tokens = Split(strSource, ChrW$(Token))
        GetTok = UBound(Tokens) + 1
        Exit Function
    End If
    '
    intPosition = InStrRev(strPosition, "-")
    intTemp = Len(strPosition) - 1
    '
    If intPosition > 0 Then
        If intPosition - intTemp = 1 Then
            'ie its at the end of the string
            intBool = True
            intRev = False
            intPos = CLng(Replace$(strPosition, "-", vbNullString)) - 1
        Else
            'must be at the start so set single token only
            intBool = False
            intRev = True
            intPos = CLng(Replace$(strPosition, "-", vbNullString)) - 1
        End If
    Else
        'just go from start
        intBool = False
        intRev = False
        intPos = CLng(strPosition) - 1
    End If
    'ok, we have our token positions lets do the dirty work
    'first split the tokens
    Tokens = Split(strSource, ChrW$(Token))
    '
    'if position is #- go from position to end
    'also check NumTokens
    If intBool = True Then
        If intRev = False Then
            If NumTokens = 0 Then
                NumTokens = UBound(Tokens)
            Else
                NumTokens = NumTokens - 1
            End If
            NumOfToks = 0
            lLower = LBound(Tokens)
            lUpper = UBound(Tokens)
            '
            For i = lLower To lUpper
                If i >= intPos Then
                    If NumOfToks <= NumTokens Then
                        NumOfToks = NumOfToks + 1
                        strToken = strToken & Tokens(i) & ChrW$(Token)
                    End If
                End If
            Next i
        End If
    End If
    'if position is -# go from beginning to position
    If intBool = False Then
        If intRev = True Then
            lLower = LBound(Tokens)
            lUpper = UBound(Tokens)
            '
            For i = lLower To lUpper
                If i <= intPos Then
                    strToken = strToken & Tokens(i) & ChrW$(Token)
                End If
            Next i
        End If
    End If
    If intBool = False Then
        If intRev = False Then
            'just return the token
            strToken = Tokens(intPos)
        End If
    End If
    'trim the end of string
    If Right$(strToken, 1) = ChrW$(Token) Then strToken = Left$(strToken, Len(strToken) - 1)
    'return it
    GetTok = strToken
End Function

'IsTok("text with delimiters", "comparator string", Character (ie: 44))
Public Function IsTok(ByVal strSource As String, ByVal strCompare As String, ByVal Token As Integer) As Boolean
    'eg: IsTok("this,is,a,test", "test", 44) = True
    'compares a source string to see if the occurance of
    'the string exists
    On Error Resume Next
    Dim Tokens() As String, i As Integer, lLower As Long, lUpper As Long
    'ok first split the string into an array
    Tokens = Split(strSource, ChrW$(Token))
    'now do some matching
    lLower = LBound(Tokens)
    lUpper = UBound(Tokens)
    For i = lLower To lUpper
        If LCase$(Tokens(i)) = LCase$(strCompare) Then
            IsTok = True
            Exit Function
        End If
    Next i
    IsTok = False
End Function

Public Function FindTok(ByVal strSource As String, ByVal strCompare As String, ByVal Occurance As Integer, ByVal Token As Integer) As Integer
    'finds and matches a token in a string of text and returns
    'its position number as an integer
    'use 0 as the 'Occurances' delimter to return the total
    'number of times the same string occurs in the source and
    '1 to return the first occurance token position
    On Error Resume Next
    Dim Tokens() As String, i As Integer, Tok As Integer, lLower As Long, lUpper As Long
    'ok, first we have to split the string into an array
    Tokens = Split(strSource, ChrW$(Token))
    Tok = 0
    'now do some matching
    lLower = LBound(Tokens)
    lUpper = UBound(Tokens)
    For i = lLower To lUpper
        If LCase$(Tokens(i)) = LCase$(strCompare) Then
            If Occurance = 0 Then
                Tok = Tok + 1
            Else
                Tok = i + 1
                Exit For
            End If
        End If
    Next i
    FindTok = Tok
End Function

Public Function AddTok(ByVal strSource As String, ByVal strAddString As String, ByVal Token As Integer) As String
    On Error Resume Next
    Dim strTemp As String
    'add the token
    strTemp = strSource & ChrW$(Token) & strAddString
    'trim the token at the front of the string
    If Left$(strTemp, 1) = ChrW$(Token) Then
        strTemp = Mid$(strTemp, 2)
    End If
    AddTok = strTemp
End Function

Public Function DelTok(ByVal strSource As String, ByVal Position As Integer, ByVal Token As Integer) As String
    'first we split the source into an array
    'then loop through looking for the position
    On Error Resume Next
    Dim Tokens() As String, i As Integer, strTemp As String, Tok As Integer, lLower As Long, lUpper As Long
    Tok = Position - 1
    '
    Tokens = Split(strSource, ChrW$(Token))
    'remove the token
    lLower = LBound(Tokens)
    lUpper = UBound(Tokens)
    For i = lLower To lUpper
        If i <> Tok Then
            strTemp = strTemp & Tokens(i) & ChrW$(Token)
        End If
    Next i
    'trim the token off the end of the string
    If Right$(strTemp, 1) = ChrW$(Token) Then strTemp = Left$(strTemp, Len(strTemp) - 1)
    DelTok = strTemp
End Function

'misc functions, InsTok (insert token at position), RepTok
'(replace token at position), PutTok (overwrites a token)
Public Function RepTok(ByVal strSource As String, ByVal strNewToken As String, ByVal Position As Integer, ByVal Token As Integer) As String
    'first split the tokens in to an array
    On Error Resume Next
    Dim Tokens() As String, i As Integer, Tok As Integer, TokTotal As Integer, strTemp As String
    Dim lLower As Long, lUpper As Long
    '
    Tok = Position - 1
    'get total number of tokens already
    TokTotal = GetTok(strSource, "0", Token) - 1
    Tokens = Split(strSource, ChrW$(Token))
    'now to replace, if the token position is out of range
    'then simply add the token to the end
    If Tok <= TokTotal Then
        lLower = LBound(Tokens)
        lUpper = UBound(Tokens)
        For i = lLower To lUpper
            If i <> Tok Then
                strTemp = strTemp & Tokens(i) & ChrW$(Token)
            Else
                'now insert the new token
                strTemp = strTemp & strNewToken & ChrW$(Token)
            End If
        Next i
    Else
        'if the token is out of range add it to the end
        strTemp = strSource & ChrW$(Token) & strNewToken
    End If
    'trim the token off the end of string
    If Right$(strTemp, 1) = ChrW$(Token) Then strTemp = Left$(strTemp, Len(strTemp) - 1)
    RepTok = strTemp
End Function

'InsTok doesn't overwrite a token but merly inserts it at
'the position
Public Function InsTok(ByVal strSource As String, ByVal strNewToken As String, ByVal Position As Integer, ByVal Token As Integer) As String
    'first split the tokens
    On Error Resume Next
    Dim Tokens() As String, i As Integer, Tok As Integer, TokTotal As Integer, strTemp As String
    Dim lLower As Long, lUpper As Long
    '
    Tokens = Split(strSource, ChrW$(Token))
    Tok = Position - 1
    TokTotal = GetTok(strSource, "0", Token) - 1
    'insert the token at position or at end if position is
    'out of range
    If Tok <= TokTotal Then
        lLower = LBound(Tokens)
        lUpper = UBound(Tokens)
        For i = lLower To lUpper
            If i <> Tok Then
                strTemp = strTemp & Tokens(i) & ChrW$(Token)
            Else
                strTemp = strTemp & strNewToken & ChrW$(Token) & Tokens(i) & ChrW$(Token)
            End If
        Next i
    Else
        'add the token at the end if its out of range
        strTemp = strSource & ChrW$(Token) & strNewToken
    End If
    'trim the token off the end of string
    If Right$(strTemp, 1) = ChrW$(Token) Then strTemp = Left$(strTemp, Len(strTemp) - 1)
    InsTok = strTemp
End Function

'PutTok overwrites a token at the specified position
Public Function PutTok(ByVal strSource As String, ByVal strNewToken As String, ByVal Position As Integer, ByVal Token As Integer) As String
    'first split the tokens
    On Error Resume Next
    Dim Tokens() As String, i As Integer, Tok As Integer, strTemp As String, lLower As Long, lUpper As Long
    '
    Tokens = Split(strSource, ChrW$(Token))
    Tok = Position - 1
    'insert the token at position or at end if position is
    'out of range
    lLower = LBound(Tokens)
    lUpper = UBound(Tokens)
    For i = lLower To lUpper
        If i <> Tok Then
            strTemp = strTemp & Tokens(i) & ChrW$(Token)
        Else
            strTemp = strTemp & strNewToken & ChrW$(Token)
        End If
    Next i
    'trim the token off the end of string
    If Right$(strTemp, 1) = ChrW$(Token) Then strTemp = Left$(strTemp, Len(strTemp) - 1)
    PutTok = strTemp
End Function
