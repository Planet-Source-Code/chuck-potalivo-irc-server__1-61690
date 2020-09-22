VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMulticast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum MulticastScopes
  DebugArea
  LAN
  WAN
  CAN
  Internet
  TimeZone
  Hemisphere
  World
End Enum

Private lngPort As Long

Private strMulticastGroup As String
Private strCurrentMulticastGroup As String

Private WithEvents sckMulticastSend As clsSocket
Attribute sckMulticastSend.VB_VarHelpID = -1
Private WithEvents sckMulticastRecieve As clsSocket
Attribute sckMulticastRecieve.VB_VarHelpID = -1

Private Declare Function intsetsockopt Lib "wsock32" Alias "setsockopt" (ByVal s As Integer, ByVal level As Integer, ByVal optname As Integer, ByRef optval As Any, ByVal optlen As Integer) As Integer

Public Event OnDataArrival(ByVal bytesTotal As Long)
Public Event OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

Public LocalIP As String

Public Ttl As Long

Private Const cbolDefaultLoopback As Boolean = True
Private bolLoopback As Boolean

Public DebugMode As Boolean

Public Property Let Loopback(bolValue As Boolean)
  bolLoopback = bolValue
  
  Dim lngLoopback As Long
  lngLoopback = IIf(bolLoopback, 1, 0)
  
  Dim lngReturn As Long
  lngReturn = intsetsockopt(sckMulticastSend.SocketHandle, IPPROTO_IP, IP_MULTICAST_LOOP, lngLoopback, Len(lngLoopback))
  
  DoEvents
  
  lngReturn = intsetsockopt(sckMulticastRecieve.SocketHandle, IPPROTO_IP, IP_MULTICAST_LOOP, lngLoopback, Len(lngLoopback))
End Property

Public Property Get Loopback() As Boolean
  Loopback = bolLoopback
End Property

Public Property Let Scope(scpValue As MulticastScopes)
  Ttl = 2 + (Scope * 32)
End Property

Public Property Get Scope() As MulticastScopes
  Scope = (Ttl - 2) / 32
End Property

Public Property Get RemoteHostIP() As String
  RemoteHostIP = sckMulticastRecieve.RemoteHostIP
End Property

Public Property Get RemotePort() As Long
  RemotePort = sckMulticastRecieve.RemotePort
End Property

Public Property Get Port() As Long
  Port = lngPort
End Property

Public Property Let Port(lngValue As Long)
  lngPort = lngValue
End Property

Public Property Get MulticastGroup() As String
  MulticastGroup = strMulticastGroup
End Property

Public Property Let MulticastGroup(strValue As String)
  strMulticastGroup = strValue
End Property

Public Function Connect()
  Dim lngReturn As Long
  Dim ipmReq As ipm_req
  
'  dbgHelper.PrintConsole "    Performing Multicast Connect"
  
  If (strCurrentMulticastGroup <> strMulticastGroup) And (Len(strCurrentMulticastGroup) > 0) Then
    ipmReq.ipm_multiaddr = inet_addr(strCurrentMulticastGroup)
    ipmReq.ipm_interface = 0
    lngReturn = intsetsockopt(sckMulticastRecieve.SocketHandle, IPPROTO_IP, IP_DROP_MEMBERSHIP, ipmReq, Len(ipmReq))
  
    If Not (lngReturn = SOCKET_ERROR) Then
    '  dbgHelper.PrintConsole "        Drop Membership: " & strCurrentMulticastGroup
    End If
  End If
  
  If strCurrentMulticastGroup <> strMulticastGroup Then
    sckMulticastRecieve.CloseSocket
    sckMulticastRecieve.Protocol = sckUDPProtocol
    sckMulticastRecieve.Bind lngPort
  
    ipmReq.ipm_multiaddr = inet_addr(strMulticastGroup)
    ipmReq.ipm_interface = 0
    lngReturn = intsetsockopt(sckMulticastRecieve.SocketHandle, IPPROTO_IP, IP_ADD_MEMBERSHIP, ipmReq, Len(ipmReq))

  '  dbgHelper.PrintConsole "        Add Membership: " & strMulticastGroup & ":" & lngPort
    If lngReturn = SOCKET_ERROR Then
    '  dbgHelper.PrintConsole "            Error: " & GetErrorDescription(Err.LastDllError)
    End If
  End If

  Dim lngLoopback As Long
  lngLoopback = IIf(bolLoopback, 1, 0)

  lngReturn = intsetsockopt(sckMulticastSend.SocketHandle, IPPROTO_IP, IP_MULTICAST_LOOP, lngLoopback, Len(lngLoopback))
  lngReturn = intsetsockopt(sckMulticastRecieve.SocketHandle, IPPROTO_IP, IP_MULTICAST_LOOP, lngLoopback, Len(lngLoopback))

'  dbgHelper.PrintConsole "        Loopback: " & bolLoopback
  If lngReturn = SOCKET_ERROR Then
  '  dbgHelper.PrintConsole "            Error: " & GetErrorDescription(Err.LastDllError)
  End If


  If DebugMode Then
  '  dbgHelper.PrintConsole "        Socket Options:"
  '  dbgHelper.PrintConsole "            SO_DEBUG: " & vbTab & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_DEBUG))
  '  dbgHelper.PrintConsole "        *   SO_ACCEPTCONN: " & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_ACCEPTCONN))
  '  dbgHelper.PrintConsole "            SO_REUSEADDR: " & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_REUSEADDR))
  '  dbgHelper.PrintConsole "        *   SO_KEEPALIVE: " & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_KEEPALIVE))
  '  dbgHelper.PrintConsole "            SO_DONTROUTE: " & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_DONTROUTE))
  '  dbgHelper.PrintConsole "            SO_BROADCAST: " & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_BROADCAST))
  '  dbgHelper.PrintConsole "        *   SO_USELOOPBACK: " & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_USELOOPBACK))
  '  dbgHelper.PrintConsole "        *   SO_LINGER: " & vbTab & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_LINGER))
  '  dbgHelper.PrintConsole "        *   SO_OOBINLINE: " & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_OOBINLINE))
  '  dbgHelper.PrintConsole "        *   SO_DONTLINGER: " & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_DONTLINGER))
  '  dbgHelper.PrintConsole "            SO_EXCLUSIVEADDRUSE: " & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_EXCLUSIVEADDRUSE))
  '  dbgHelper.PrintConsole "            SO_SNDBUF: " & vbTab & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_SNDBUF))
  '  dbgHelper.PrintConsole "            SO_RCVBUF: " & vbTab & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_RCVBUF))
  '  dbgHelper.PrintConsole "            SO_ERROR: " & vbTab & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_ERROR))
  '  dbgHelper.PrintConsole "            SO_TYPE: " & vbTab & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, SOL_SOCKET, SO_TYPE))
  '  dbgHelper.PrintConsole "        TCP Options:" & vbTab & vbTab & vbTab & vbTab
  '  dbgHelper.PrintConsole "        *   TCP_NODELAY: " & vbTab & vbTab & vbTab & CStr(GetSocketOption(sckMulticastRecieve.SocketHandle, IPPROTO_TCP, TCP_NODELAY))
  End If
  
  strCurrentMulticastGroup = strMulticastGroup
End Function

Private Sub Class_Initialize()
  Set sckMulticastRecieve = New clsSocket
  
  Set sckMulticastSend = New clsSocket
  
  LocalIP = GetDefaultIP
  Scope = World

  bolLoopback = cbolDefaultLoopback
End Sub

Private Sub Class_Terminate()
  sckMulticastRecieve.CloseSocket
  Set sckMulticastRecieve = Nothing
  
  sckMulticastSend.CloseSocket
  Set sckMulticastSend = Nothing
End Sub

Private Sub sckMulticastRecieve_OnDataArrival(ByVal bytesTotal As Long)
  RaiseEvent OnDataArrival(bytesTotal)
End Sub

Public Sub GetData(ByRef varData, Optional varType, Optional maxLen)
  sckMulticastRecieve.GetData varData, varType, maxLen
End Sub

Public Sub SendData(varData)
  If DebugMode Then
  '  dbgHelper.PrintConsole "     Sending Data: " & varData
  '  dbgHelper.PrintConsole "         Port: " & lngPort
  '  dbgHelper.PrintConsole "         RemoteHost: " & strMulticastGroup
  End If
  
  Set sckMulticastSend = New clsSocket

  sckMulticastSend.Protocol = sckUDPProtocol
  sckMulticastSend.RemotePort = lngPort
  sckMulticastSend.RemoteHost = strMulticastGroup

  Dim lngReturn As Long
  lngReturn = intsetsockopt(sckMulticastSend.SocketHandle, IPPROTO_IP, IP_MULTICAST_TTL, Ttl, Len(Ttl))
  
  sckMulticastSend.SendData varData
End Sub

Private Sub sckMulticastRecieve_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  RaiseEvent OnError(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub sckMulticastSend_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  RaiseEvent OnError(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Public Sub CloseMulticast()
  On Error Resume Next

  Dim lngReturn As Long
  Dim ipmReq As ipm_req
  
  If Len(strCurrentMulticastGroup) > 0 Then
    ipmReq.ipm_multiaddr = inet_addr(strCurrentMulticastGroup)
    ipmReq.ipm_interface = 0
    lngReturn = intsetsockopt(sckMulticastRecieve.SocketHandle, IPPROTO_IP, IP_DROP_MEMBERSHIP, ipmReq, Len(ipmReq))
    
    strCurrentMulticastGroup = ""
  End If
  
  sckMulticastSend.CloseSocket
  sckMulticastRecieve.CloseSocket
End Sub
