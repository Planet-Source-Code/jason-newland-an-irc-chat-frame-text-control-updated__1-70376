Attribute VB_Name = "mMain"

Option Explicit

Public b_IsNT As Boolean

Public Sub Main()
    On Error Resume Next
    'Check that our OS is NT so we can use unicode
    b_IsNT = IsNT
    '
    LoadScheme
    frmTest.Show vbModeless
End Sub

Private Function IsNT() As Boolean
    On Error Resume Next
    Dim udtVer As OSVERSIONINFO
    '
    udtVer.dwOSVersionInfoSize = Len(udtVer)
    If GetVersionEx(udtVer) Then
        If udtVer.dwMajorVersion >= 5 Then
            IsNT = True
            Exit Function
        End If
    End If
    IsNT = False
End Function
