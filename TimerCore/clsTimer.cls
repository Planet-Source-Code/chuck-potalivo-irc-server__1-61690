VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private bolEnabled As Boolean
Private lngInterval As Long

Private lngEventID As Long

Public Tag As Variant
Public Event Timer()

Public Property Get Version() As String
  Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Property Let Enabled(bolValue As Boolean)
  If bolValue And Not bolEnabled Then
    Create
  ElseIf bolValue And bolEnabled Then
    ' It's already on
  ElseIf Not bolValue And bolEnabled Then
    Kill
  ElseIf Not bolValue And Not bolEnabled Then
    ' It's already off
  End If
  
  bolEnabled = bolValue
End Property

Property Get Enabled() As Boolean
  Enabled = bolEnabled
End Property

Property Let Interval(lngValue As Long)
  Dim lngOldInterval As Long

  lngOldInterval = lngInterval
  lngInterval = lngValue
  
  If bolEnabled And (lngOldInterval <> lngInterval) Then
    Kill
    DoEvents
    Create
  End If
End Property

Property Get Interval() As Long
  Interval = lngInterval
End Property

Private Sub Class_Initialize()
  If colTimers Is Nothing Then
    Set colTimers = New Collection
  End If
End Sub

Private Sub Class_Terminate()
  If bolEnabled Then
    Kill
  End If
End Sub

Friend Sub Timer()
  RaiseEvent Timer
End Sub

Public Sub Terminate()
  Kill
End Sub

Private Sub Kill()
  Dim lngReturn As Long
  
  lngReturn = KillTimer(0, lngEventID)
  
  colTimers.Remove "ID" & lngEventID
  
  bolEnabled = False
End Sub

Private Sub Create()
  Dim lngReturn As Long

  lngReturn = SetTimer(0, lngEventID, lngInterval, AddressOf TimerProc)
  lngEventID = lngReturn
  
  colTimers.Add Me, "ID" & lngEventID
  
  bolEnabled = True
End Sub
