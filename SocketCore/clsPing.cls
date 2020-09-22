VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, RequestOptions As ICMP_OPTIONS, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long 'formerly integer
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Declare Function lstrlen Lib "kernel32" _
   Alias "lstrlenA" _
  (lpString As Any) As Long
  
Public DataToSend As String

Public Function Ping(sAddress As String) As Long
  'If Ping succeeds :
  '.RoundTripTime = time in ms for the ping to complete,
  '.Data is the data returned (NULL terminated)
  '.Address is the Ip address that actually replied
  '.DataSize is the size of the string in .Data
  '.Status will be 0
  '
  'If Ping fails .Status will be the error code
   
   Dim ECHO As ICMP_ECHO_REPLY
   Dim hPort As Long
   Dim dwAddress As Long
   Dim PING_TIMEOUT As Long
   Dim ipo As ICMP_OPTIONS
   
  'convert the address into a long representation
   dwAddress = inet_addr(sAddress)
   
  'if a valid address..
   If dwAddress <> INADDR_NONE Then
     'open a port
      hPort = IcmpCreateFile()
      
     'and if successful,
      If hPort Then
        'ping it.
'         Call IcmpSendEcho(hPort, dwAddress, DataToSend, Len(DataToSend), ipo, ECHO, Len(ECHO), PING_TIMEOUT)

        'return the status as ping succes and close
         Ping = ECHO.RoundTripTime
         Call IcmpCloseHandle(hPort)
      End If
   Else:
        'the address format was probably invalid
         Ping = INADDR_NONE
   End If
End Function

Public Function HopCount(sAddress As String) As Long
   Dim ECHO As ICMP_ECHO_REPLY
   Dim hPort As Long
   Dim dwAddress As Long
   Dim PING_TIMEOUT As Long
   Dim ipo As ICMP_OPTIONS
   
  'convert the address into a long representation
   dwAddress = inet_addr(sAddress)
   
  'if a valid address..
   If dwAddress <> INADDR_NONE Then
     'open a port
      hPort = IcmpCreateFile()
      
     'and if successful,
      If hPort Then
         Dim lTTL As Long
         Dim sHost As String
         
        'increment the ttl until we hit the destination host
         For lTTL = 1 To 255
           ipo.Ttl = lTTL
      
          'ping it.
           Call IcmpSendEcho(hPort, dwAddress, DataToSend, Len(DataToSend), ipo, ECHO, Len(ECHO), PING_TIMEOUT)
           
           sHost = GetIPFromAddress(ECHO.Address)
           
           If sHost = sAddress Then
             HopCount = lTTL
           
             Exit For
           End If
         Next lTTL

        'return the status as ping succes and close
         Call IcmpCloseHandle(hPort)
      End If
   Else:
        'the address format was probably invalid
         HopCount = 0
   End If
End Function

Private Function SocketsInitialize() As Boolean
   Dim WSAD As WSAData
   
   SocketsInitialize = (WSAStartup(&H101, WSAD) = 0)
End Function


Private Sub SocketsCleanup()
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
End Sub

Public Function Hostname(ByVal sAddress As String) As String
   Dim ptrHosent As Long
   Dim hAddress As Long
   Dim nbytes As Long
   
   If SocketsInitialize() Then

     'convert string address to long
      hAddress = inet_addr(sAddress)
      
      If hAddress <> SOCKET_ERROR Then
         
        'obtain a pointer to the HOSTENT structure
        'that contains the name and address
        'corresponding to the given network address.
         ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
   
         If ptrHosent <> 0 Then
         
           'convert address and
           'get resolved hostname
            CopyMemory ptrHosent, ByVal ptrHosent, 4
            nbytes = lstrlen(ByVal ptrHosent)
         
            If nbytes > 0 Then
               sAddress = Space$(nbytes)
               CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
               Hostname = sAddress
            End If
         
         Else:
           Hostname = sAddress
         End If 'If ptrHosent
      
      SocketsCleanup
      
      Else: 'MsgBox "String passed is an invalid IP."
        Hostname = sAddress
      End If 'If hAddress
   
   Else: 'MsgBox "Sockets failed to initialize."
     Hostname = sAddress
   End If  'If SocketsInitialize
End Function

Private Sub Class_Initialize()
  DataToSend = Space(32)
End Sub
