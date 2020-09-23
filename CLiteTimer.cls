VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLiteTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mlngInterval As Long ' The timer interval.
Private mlngTimerID As Long

Dim tmrTag As String

Public Event Timer()

Public Property Let Tag(ByVal strTag As String)
    On Error Resume Next
    tmrTag = strTag
End Property

Public Property Get Tag() As String
    On Error Resume Next
    Tag = tmrTag
End Property

Public Property Let Enabled(ByVal blnEnabled As Boolean)
    On Error GoTo ErrorHandler
    ' Is the timer being started?
    If blnEnabled Then
        ' Is the timer currently running?
        If Not MLiteTimer.TimerRunning(mlngTimerID) Then
            ' Start the timer.
            mlngTimerID = MLiteTimer.StartTimer(Me, mlngInterval, mlngTimerID)
        End If
    Else
        ' Stop the timer.
        MLiteTimer.StopTimer mlngTimerID
    End If
    Exit Property
ErrorHandler:
    'Debug.Print "Enabled Let Error " & Err.Number & ": " & Err.Description
End Property

Public Property Get Enabled() As Boolean
    On Error GoTo ErrorHandler
    Enabled = MLiteTimer.TimerRunning(mlngTimerID)
    Exit Property
ErrorHandler:
    'Debug.Print "Enabled Get Error " & Err.Number & ": " & Err.Description
End Property

Public Property Let Interval(ByVal lngInterval As Long)
    On Error GoTo ErrorHandler
    ' Check the interval value.
    If lngInterval > 0 Then
        ' Check to see if the timer is currently running.
        If MLiteTimer.TimerRunning(mlngTimerID) Then
            ' Stop the timer.
            MLiteTimer.StopTimer mlngTimerID
            ' Restart the timer with the new interval.
            mlngTimerID = MLiteTimer.StartTimer(Me, lngInterval, mlngTimerID)
        End If
        ' Store the new interval value.
        mlngInterval = lngInterval
    Else
        mlngInterval = 0
        ' Stop the timer.
        MLiteTimer.StopTimer mlngTimerID
    End If
    Exit Property
ErrorHandler:
    'Debug.Print "Interval Let Error " & Err.Number & ": " & Err.Description
End Property

Public Property Get Interval() As Long
    On Error GoTo ErrorHandler
    Interval = mlngInterval
    Exit Property
ErrorHandler:
    'Debug.Print "Interval Get Error " & Err.Number & ": " & Err.Description
End Property

Friend Sub TimerCallBack(ByVal lngTimerID As Long)
    On Error GoTo ErrorHandler
    ' Raise the event if the timer id's match.
    If mlngTimerID = lngTimerID Then
        RaiseEvent Timer
    End If
    Exit Sub
ErrorHandler:
    'Debug.Print "TimerCallBack Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    'Debug.Print "Creating Timer Object"
    ' Set a default timer interval.
    mlngInterval = 1000
    Exit Sub
ErrorHandler:
    'Debug.Print "Class_Initialize Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    ' Stop the timer if it is running.
    MLiteTimer.StopTimer mlngTimerID
End Sub
