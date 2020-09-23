Attribute VB_Name = "mUTF8"
'-----------------------------------------------------------------------------
'   InCodePage
'       Lots of utility functions for deteremining which codepage and LCID
'       to use for an arbitrary string, very helpful adjunct in multinational
'       applications where you may not know the language.
'
'   Author: MichKa
'
'   Copyright © 1999 Trigeminal Software, Inc. All Rights Reserved.
'
'   You may use this code freely in your applications, you just can't
'   put it in a book or article and call it your own.
'-----------------------------------------------------------------------------
Option Compare Binary
Option Explicit

Public Const CP_ACP = 0
Public Const CP_NONE = 0
Public Const CP_OEMCP = 1
Public Const CP_WINANSI = 1004
Public Const CP_WINUNICODE = 1200
Public Const MB_PRECOMPOSED = &H1
Public Const MB_COMPOSITE = &H2
Public Const MB_USEGLYPHCHARS = &H4

' //Arabic
Public Const CP_AWIN = 101 ' //Bidi Windows codepage
Public Const CP_709 = 102  ' //MS-DOS Arabic Support CP 709
Public Const CP_720 = 103  ' //MS-DOS Arabic Support CP 720
Public Const CP_A708 = 104 ' //ASMO 708
Public Const CP_A449 = 105 ' //ASMO 449+
Public Const CP_TARB = 106 ' //MS Transparent Arabic
Public Const CP_NAE = 107  ' //Nafitha Enhanced Arabic Char Set
Public Const CP_V4 = 108   ' //Nafitha v 4.0
Public Const CP_MA2 = 109  ' //Mussaed Al Arabi (MA/2) CP 786
Public Const CP_I864 = 110 ' //IBM Arabic Supplement CP 864
Public Const CP_A437 = 111 ' //Ansi 437 codepage
Public Const CP_AMAC = 112 ' //Macintosh Cod Page
' //Hebrew
Public Const CP_HWIN = 201 ' //Bidi Windows codepage
Public Const CP_862I = 202 ' //IBM Hebrew Supplement CP 862
Public Const CP_7BIT = 203 ' //IBM Hebrew Supplement CP 862 Folded
Public Const CP_ISO = 204   ' //ISO Hebrew 8859-8 Character Set
Public Const CP_H437 = 205 ' //Ansi 437 codepage
Public Const CP_HMAC = 206 ' //Macintosh Cod Page

' /*************************************************************************
'    Code Pages
' *************************************************************************/
Public Const CP_OEM_437 = 437
Public Const CP_ARABICDOS = 708
Public Const CP_DOS720 = 720
Public Const CP_IBM850 = 850
Public Const CP_IBM852 = 852
Public Const CP_DOS862 = 862
Public Const CP_IBM866 = 866
Public Const CP_THAI = 874
Public Const CP_JAPAN = 932
Public Const CP_CHINA = 936
Public Const CP_KOREA = 949
Public Const CP_TAIWAN = 950
Public Const CP_EASTEUROPE = 1250
Public Const CP_RUSSIAN = 1251
Public Const CP_WESTEUROPE = 1252
Public Const CP_GREEK = 1253
Public Const CP_TURKISH = 1254
Public Const CP_HEBREW = 1255
Public Const CP_ARABIC = 1256
Public Const CP_BALTIC = 1257
Public Const CP_VIETNAMESE = 1258
Public Const CP_ASCII = 20127
Public Const CP_RUSSIANKOI8R = 20866
Public Const CP_RUSSIANKOI8U = 21866
Public Const CP_ISOLATIN1 = 28591
Public Const CP_ISOEASTEUROPE = 28592
Public Const CP_ISOTURKISH = 28593
Public Const CP_ISOBALTIC = 28594
Public Const CP_ISORUSSIAN = 28595
Public Const CP_ISOARABIC = 28596
Public Const CP_ISOGREEK = 28597
Public Const CP_ISOHEBREW = 28598
Public Const CP_ISOTURKISH2 = 28599
Public Const CP_ISOLATIN9 = 28605
Public Const CP_HEBREWLOG = 38598
Public Const CP_USER = 50000
Public Const CP_AUTOALL = 50001
Public Const CP_JAPANNHK = 50220
Public Const CP_JAPANESC = 50221
Public Const CP_JAPANSIO = 50222
Public Const CP_KOREAISO = 50225
Public Const CP_TAIWANISO = 50227
Public Const CP_CHINAISO = 50229
Public Const CP_AUTOJAPAN = 50932
Public Const CP_AUTOCHINA = 50936
Public Const CP_AUTOKOREA = 50949
Public Const CP_AUTOTAIWAN = 50950
Public Const CP_AUTORUSSIAN = 51251
Public Const CP_AUTOGREEK = 51253
Public Const CP_AUTOARABIC = 51256
Public Const CP_JAPANEUC = 51932
Public Const CP_CHINAEUC = 51936
Public Const CP_KOREAEUC = 51949
Public Const CP_TAIWANEUC = 51950
Public Const CP_CHINAHZ = 52936
Public Const CP_MAC_ROMAN = 10000
Public Const CP_MAC_JAPAN = 10001
Public Const CP_MAC_ARABIC = 10004
Public Const CP_MAC_GREEK = 10006
Public Const CP_MAC_CYRILLIC = 10007
Public Const CP_MAC_LATIN2 = 10029
Public Const CP_MAC_TURKISH = 10081

#If Mac Then
    Public Const CP_DEFAULT = CP_MACCP
#Else
    Public Const CP_DEFAULT = CP_ACP
#End If

Public Const CP_JOHAB = 1361
Public Const CP_SYMBOL = 42
Public Const CP_UTF8 = 65001
Public Const CP_UTF7 = 65000
Public Const CP_UNICODELITTLE = 1200
Public Const CP_UNICODEBIG = 1201
Public Const CP_UNKNOWN = -1

Public Const MB_ERR_INVALID_CHARS = &H8             ' /* error for invalid chars */
Public Const WC_DEFAULTCHECK = &H100                ' /* check for default char */
Public Const WC_COMPOSITECHECK = &H200              ' /* convert composite to precomposed */
Public Const WC_DISCARDNS = &H10                    ' /* discard non-spacing chars */
Public Const WC_SEPCHARS = &H20                     ' /* generate separate chars */
Public Const WC_DEFAULTCHAR = &H40                  ' /* replace w/ default char */

Private Declare Function GetACP Lib "kernel32" () As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long

Public Function AToW(ByVal st As String, Optional ByVal cpg As Long = -1, Optional lFlags As Long = 0) As String
    ' ANSI to UNICODE conversion, via a given codepage.
    On Error Resume Next
    Dim stBuffer As String, cwch As Long, pwz As Long, pwzBuffer As Long
    If cpg = -1 Then cpg = GetACP()
    pwz = StrPtr(st)
    cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, 0&, 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer))
    AToW = Left$(stBuffer, cwch - 1)
End Function

Public Function WToA(ByVal st As String, Optional ByVal cpg As Long = -1, Optional lFlags As Long = 0) As String
    ' UNICODE to ANSI conversion, via a given codepage
    On Error Resume Next
    Dim stBuffer As String, cwch As Long, pwz As Long, pwzBuffer As Long, lpUsedDefaultChar As Long
    If cpg = -1 Then cpg = GetACP()
    pwz = StrPtr(st)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
    WToA = Left$(stBuffer, cwch - 1)
End Function

