Attribute VB_Name = "modMulticast"
Option Explicit

Public Type ipm_req
    ipm_multiaddr As Long
    ipm_interface As Long
End Type


Global Const IP_MULTICAST_IF As Long = 2
Global Const IP_MULTICAST_TTL As Long = 3
Global Const IP_MULTICAST_LOOP As Long = 4
Global Const IP_ADD_MEMBERSHIP As Long = 5
Global Const IP_DROP_MEMBERSHIP As Long = 6

Global Const TCP_NODELAY = &H1&

Public Type LINGER_STRUCT
  l_onoff As Integer
  l_linger As Integer
End Type

Public Function GetSocketOption(lSocket As Long, lLevel As Long, lOption As Long) As Long
  Dim lResult As Long       ' Result of API call.
  Dim lBuffer As Long       ' Buffer to get value into.
  Dim lBufferLen As Long    ' len of buffer.
  Dim linger As LINGER_STRUCT
  
  ' Linger requires a structure so we will get that option differently.
  If (lOption <> SO_LINGER) And (lOption <> SO_DONTLINGER) Then
    lBufferLen = LenB(lBuffer)
    lResult = getsockopt(lSocket, lLevel, lOption, lBuffer, lBufferLen)
  Else
    lBufferLen = LenB(linger)
    lResult = getsockopt(lSocket, lLevel, lOption, linger, lBufferLen)
    lBuffer = linger.l_onoff
  End If
  
  If (lResult = SOCKET_ERROR) Then
    GetSocketOption = -1
  Else
    GetSocketOption = lBuffer
  End If
End Function

