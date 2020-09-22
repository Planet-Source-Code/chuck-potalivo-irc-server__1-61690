Attribute VB_Name = "modTimer"
Option Explicit

Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Global colTimers As Collection

Public Sub TimerProc(ByVal hwnd As Long, ByVal uint1 As Long, ByVal nEventId As Long, ByVal dwParam As Long)
  On Error GoTo ErrH
  
  Dim tmrTemp As clsTimer
  
  Set tmrTemp = colTimers("ID" & nEventId)
  
  tmrTemp.Timer
  
  Exit Sub
ErrH:
  KillTimer 0, nEventId
  
  Exit Sub
End Sub
