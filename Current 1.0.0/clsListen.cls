VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents sckListen      As clsSocket
Attribute sckListen.VB_VarHelpID = -1

Private strGUID                   As String
Private strDebugGUID              As String

Public Property Get GUID() As String
  GUID = strGUID
End Property

Public Property Get IP() As String
  IP = sckListen.LocalIP
End Property

Public Property Get Port() As Long
  Port = sckListen.LocalPort
End Property

Public Function Initialize(lngPort As Long, Optional strIP As String)
  Set sckListen = New clsSocket
  sckListen.Protocol = sckTCPProtocol
  
  sckListen.Bind lngPort, strIP
  sckListen.Listen
  
  Dim strListen As String
  strListen = Space(20)
  
  LSet strListen = "SOCK " & sckListen.LocalIP

  strGUID = GenerateGUID

  strDebugGUID = Replace(strGUID, "-", "")
  strDebugGUID = Replace(strDebugGUID, "}", "")
  strDebugGUID = Replace(strDebugGUID, "{", "")
  strDebugGUID = Left(strDebugGUID, (Len(strDebugGUID) / 2) - 6)
  strDebugGUID = strListen & "|" & strDebugGUID
  
  DebugPrint "[" & strDebugGUID & " |INIT  ] IRC Server Listen"
  DebugPrint "[" & strDebugGUID & " |INIT  ]   GUID:      " & strGUID
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Hostname:  " & sckListen.LocalHostName
  DebugPrint "[" & strDebugGUID & " |INIT  ]   IP:        " & sckListen.LocalIP
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Port:      " & sckListen.LocalPort
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Thread ID: " & GetCurrentThreadId
End Function

Private Sub Class_Terminate()
  sckListen.CloseSocket
  Set sckListen = Nothing

  DebugPrint "[" & strDebugGUID & " |CLOSE ] Terminated"
End Sub

Private Sub sckListen_OnConnectionRequest(ByVal requestID As Long)
  DebugPrint "[" & strDebugGUID & " |CONREQ] RequestID: " & requestID
  
  CreateConnect requestID
End Sub
