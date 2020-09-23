Attribute VB_Name = "MLiteTimer"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)

Private Const WM_TIMER = &H113

Private mobjTimers As Collection

Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error GoTo ErrorHandler
    Dim objTimer As CLiteTimer
    ' Make sure that the message is WM_TIMER.
    If uMsg = WM_TIMER Then
        For Each objTimer In mobjTimers
            ' Execute the callback method in the class.
            objTimer.TimerCallBack idEvent
        Next objTimer
    End If
    Exit Sub
ErrorHandler:
    'Debug.Print "TimerProc Error " & Err.Number & ": " & Err.Description
End Sub

Public Function StartTimer(ByVal objTimer As CLiteTimer, ByVal lngInterval As Long, ByVal lngTimerID As Long) As Long
    On Error GoTo ErrorHandler
    Dim lngTimerPtr As Long, objTimerCopy As CLiteTimer
    ' Make sure there is a parent object.
    Debug.Assert Not (objTimer Is Nothing)
    ' Create the collection to store the timers if it hasn't been already.
    If mobjTimers Is Nothing Then
        Set mobjTimers = New Collection
    End If
    ' Check to see if the timer is already running.
    If lngTimerID = 0 Then
        ' No timer is running.
        ' Was an interval specified?
        If lngInterval > 0 Then
            ' Everything is okay.
            ' Now create the timer.
            lngTimerID = SetTimer(0, 0, lngInterval, AddressOf TimerProc)
            ' Get a pointer to the object. This enables the use of weak pointers.
            lngTimerPtr = ObjPtr(objTimer)
            ' Copy the pointer to another object. This object
            ' will be used from now on to reference the parent.
            CopyMemory objTimerCopy, lngTimerPtr, 4
            mobjTimers.Add objTimerCopy, "T" & lngTimerID
            'Debug.Print "StartTimer", StartTimer, lngInterval
        End If
    End If
    StartTimer = lngTimerID
    Exit Function
ErrorHandler:
    'Debug.Print "StartTimer Error " & Err.Number & ": " & Err.Description
End Function

Public Sub StopTimer(ByRef lngTimerID As Long)
    On Error GoTo ErrorHandler
    Dim objTimerCopy As CLiteTimer
    If Not (mobjTimers Is Nothing) Then
        ' Is the timer running?
        If TimerRunning(lngTimerID) Then
            ' The timer is running. Kill it.
            If KillTimer(0, lngTimerID) <> 0 Then
                ' Timer killed.
                ' Get a reference to the parent object.
                ' This needs to be overwritten.
                Set objTimerCopy = mobjTimers("T" & lngTimerID)
                ' Remove the parent from the collection.
                ' This will not destroy the object as we
                ' still have a reference.
                mobjTimers.Remove "T" & lngTimerID
                ' Now finally destroy the reference to the object.
                ' Do not set the object to nothing as this will
                ' decrease the refcount which will cause VB to crash.
                CopyMemory objTimerCopy, 0&, 4
                lngTimerID = 0
                If mobjTimers.Count = 0 Then
                    Set mobjTimers = Nothing
                End If
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    'Debug.Print "StopTimer Error " & Err.Number & ": " & Err.Description
End Sub

Public Property Get TimerRunning(ByVal lngTimerID As Long) As Boolean
    On Error GoTo ErrorHandler
    TimerRunning = (lngTimerID <> 0)
    Exit Property
ErrorHandler:
    'Debug.Print "TimerRunning Get Error " & Err.Number & ": " & Err.Description
End Property
