VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private aryStack()          As Variant

Private Sub Class_Initialize()
  ReDim aryStack(0)
End Sub

Private Sub Class_Terminate()
  ReDim aryStack(0)
End Sub

Public Function Add(varItem As Variant)
  ReDim Preserve aryStack(0 To (UBound(aryStack) + 1))
  
  aryStack(UBound(aryStack)) = varItem
End Function

Public Function NextItem() As Variant
  If UBound(aryStack) < 1 Then
    Exit Function
  End If

  NextItem = aryStack(1)
  
  Dim lngIndex              As Long
  
  For lngIndex = 1 To UBound(aryStack)
    aryStack(lngIndex - 1) = aryStack(lngIndex)
  Next lngIndex
  
  ReDim Preserve aryStack(0 To (UBound(aryStack) - 1))
End Function

Public Property Get Count() As Long
  Count = UBound(aryStack)
End Property
