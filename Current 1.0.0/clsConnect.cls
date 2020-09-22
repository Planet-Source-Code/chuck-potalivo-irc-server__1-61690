VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents sckConnect     As clsSocket
Attribute sckConnect.VB_VarHelpID = -1

Private WithEvents tmrConnect     As clsTimer
Attribute tmrConnect.VB_VarHelpID = -1
Private WithEvents tmrSend        As clsTimer
Attribute tmrSend.VB_VarHelpID = -1

Private stkSend                   As clsStack

Private strGUID                   As String
Private strDebugGUID              As String

Private strNick                   As String
Private strModes                  As String

Private strAway                   As String

Private strUsername               As String
Private strRealName               As String

Private strIP                     As String
Private strHostName               As String

Private lngSignon                 As Long
Private lngIdleTimer              As Long

Private bolGotIdent               As Boolean
Private bolAuthed                 As Boolean

Private bolPong                   As Boolean
Private lngPing                   As Long

Private bolConnectOk              As Boolean

Public Property Get DebugGUID() As String
  DebugGUID = strDebugGUID
End Property

Public Property Get Ping() As Long
  Ping = lngPing
End Property

Public Property Get GUID() As String
  GUID = strGUID
End Property

Public Property Get Nick() As String
  Nick = strNick
End Property

Public Property Get Username() As String
  Username = strUsername
End Property

Public Property Get hostname() As String
  hostname = strHostName
End Property

Public Property Get RemoteIP() As String
  RemoteIP = sckConnect.RemoteHostIP
End Property

Public Property Get LocalIP() As String
  LocalIP = sckConnect.LocalIP
End Property

Public Property Get RealName() As String
  RealName = strRealName
End Property

Public Property Get IdleTime() As Long
  IdleTime = Timer - lngIdleTimer
End Property

Public Property Get FullIdent() As String
  FullIdent = strNick & "!" & strUsername & "@" & strHostName
End Property

Public Property Get Away() As String
  Away = strAway
End Property

Public Property Get Mode() As String
  Mode = strModes
End Property

Public Function Initialize(lngRequestID As Long, Optional bolOk As Boolean = True)
  bolConnectOk = bolOk

  Set sckConnect = New clsSocket
  sckConnect.Protocol = sckTCPProtocol
  
  sckConnect.Accept lngRequestID
  
  Dim strRemoteHostIP As String
  strRemoteHostIP = Space(20)
  
  LSet strRemoteHostIP = "CONN " & sckConnect.RemoteHostIP
  
  strGUID = GenerateGUID
  
  strDebugGUID = Replace(strGUID, "-", "")
  strDebugGUID = Replace(strDebugGUID, "}", "")
  strDebugGUID = Replace(strDebugGUID, "{", "")
  strDebugGUID = Left(strDebugGUID, (Len(strDebugGUID) / 2) - 6)
  strDebugGUID = strRemoteHostIP & "|" & strDebugGUID
  
  strNick = ""
  bolAuthed = False
  
  lngSignon = UnixTime
  lngIdleTimer = Timer
  
  strModes = "w"
  strAway = ""
  
  lngPing = 0
  bolPong = True
  
  Set tmrConnect = New clsTimer
  tmrConnect.Interval = glngPingInterval
  tmrConnect.Enabled = True
  
  Set stkSend = New clsStack
  
  Set tmrSend = New clsTimer
  tmrSend.Interval = glngSendInterval
  tmrSend.Enabled = True
  
  DebugPrint "[" & strDebugGUID & " |INIT  ] Connect Class"
  DebugPrint "[" & strDebugGUID & " |INIT  ]   GUID:      " & strGUID
  DebugPrint "[" & strDebugGUID & " |INIT  ]   IP:        " & sckConnect.RemoteHostIP & ":" & sckConnect.RemotePort
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Thread ID: " & GetCurrentThreadId
End Function

Private Sub Class_Terminate()
  sckConnect.CloseSocket
  Set sckConnect = Nothing
  
  tmrConnect.Enabled = False
  Set tmrConnect = Nothing

  DebugPrint "[" & strDebugGUID & " |CLOSE ] Terminated"
End Sub

Public Function SendData(strData As String)
  If sckConnect.State <> sckConnected Then
    Exit Function
  End If
  
  stkSend.Add strData & vbCrLf
End Function

Private Sub sckConnect_OnDataArrival(ByVal bytesTotal As Long)
  If Not bolConnectOk Then
    Exit Sub
  End If

  Dim strData     As String
  
  sckConnect.GetData strData, vbString
  
  strData = Replace(strData, vbCrLf, vbLf)
  strData = Left(strData, Len(strData) - 1)
  
  Dim aryCmdLns() As String
  aryCmdLns = Split(strData, vbLf)
  
  Dim lngCmdLnIndex As Long
  
  Dim aryCmds()    As String
  
  Dim strCmdLn    As String
  Dim strCmd      As String
  
  For lngCmdLnIndex = LBound(aryCmdLns) To UBound(aryCmdLns)
    strCmdLn = aryCmdLns(lngCmdLnIndex)
    
    aryCmds = SmartSplit(strCmdLn, " ")
    ReDim Preserve aryCmds(0 To 100)
    
    strCmd = aryCmds(0)
    strCmd = UCase(strCmd)
    
    If gbolShowRcvdData Then
      DebugPrint "[" & strDebugGUID & " |RCVD  ] " & strCmdLn
    End If
    
    If Not bolAuthed And Not ((strCmd Like "NICK") Or (strCmd Like "USER")) Then
      ' ERR_NOTREGISTERED
      SendData ":" & gstrServerName & " 451 :You have not registered"
      
      GoTo NextCmd
    End If
    
    Select Case strCmd
      Case "MODE":
        Select Case True
          Case (Len(aryCmds(1)) = 0):
            ' ERR_NEEDMOREPARAMS
            SendData ":" & gstrServerName & " 461 MODE :Not enough parameters"
          Case IsChannel(aryCmds(1)):
            Dim chnMode As clsChannel
            Set chnMode = GetChannel(aryCmds(1))
            
            If (Len(aryCmds(2)) > 0) And (Len(aryCmds(3)) = 0) Then
              Select Case LCase(aryCmds(2))
                Case "+b", "b":
                  chnMode.SendBanList strNick
                Case "+m":
                  chnMode.Moderated(Me.FullIdent) = True
                Case "-m":
                  chnMode.Moderated(Me.FullIdent) = False
                Case Else:
                  ' ERR_UMODEUNKNOWNFLAG
                  SendData ":" & gstrServerName & " 501 :Unknown MODE flag"
              End Select
            
              ' RPL_CHANNELMODEIS
              'SendData ":" & gstrServerName & " 324 " & chnMode.Name & " " & LCase(aryCmds(3)) & " " & strModeParams
            ElseIf (Len(aryCmds(2)) > 0) And (Len(aryCmds(3)) > 0) Then
              Select Case LCase(aryCmds(2))
                Case "+b":
                  chnMode.BanNick strNick, aryCmds(3)
                Case "-b":
                Case Else:
                  ' ERR_UMODEUNKNOWNFLAG
                  SendData ":" & gstrServerName & " 501 :Unknown MODE flag"
              End Select
            ElseIf (Len(aryCmds(2)) = 0) Then
              SendData ":" & gstrServerName & " MODE " & chnMode.Name & " +" & chnMode.Mode
            End If
          Case IsConnect(aryCmds(1)):
            Dim conMode As clsConnect
            Set conMode = GetConnect(aryCmds(1))
            
            If Len(aryCmds(2)) = 0 Then
              ' RPL_UMODEIS
              SendData ":" & gstrServerName & " 221 " & conMode.Nick & " +" & conMode.Mode
            ElseIf Not (aryCmds(2) Like strNick) Then
              ' ERR_USERSDONTMATCH
              SendData ":" & gstrServerName & " 502 :Cannot change mode for other users"
            Else
              ' CHANGE THE MODE!
            End If
          Case Else:
            ' ERR_NOSUCHNICK
            SendData ":" & gstrServerName & " 401 " & aryCmds(1) & " :No such nick/channel"
        End Select
      Case "NICK":
        ChangeNick aryCmds(1)
      Case "USER":
        SetUserInfo aryCmds(1), aryCmds(2), aryCmds(3), aryCmds(4)
      Case "PONG":
        bolPong = True
      
        Dim dblTimer As Double
        dblTimer = Val(aryCmds(1))
        
        lngPing = (Timer - dblTimer) * 1000
        
        If lngPing < 0 Then
          lngPing = 0
        End If
        
        DebugPrint "[" & strDebugGUID & " |PING  ] " & strNick & " | " & lngPing & " ms"
      Case "JOIN":
        JoinChannel aryCmds(1), strNick
      Case "PART":
        PartChannel aryCmds(1), strNick, aryCmds(2)
      Case "KICK":
        KickNick strNick, aryCmds(1), aryCmds(2), aryCmds(3)
      Case "TOPIC":
        SetChanTopic aryCmds(1), strNick, aryCmds(2)
      Case "QUIT":
        Quit aryCmds(1)
      Case "AWAY":
        strAway = aryCmds(1)
        
        DebugPrint "[" & strDebugGUID & " |AWAY  ] " & strAway
        
        If Len(strAway) = 0 Then
          ' RPL_UNAWAY
          SendData ":" & gstrServerName & " 305 :You are no longer marked as being away"
        Else
          ' RPL_NOWAWAY
          SendData ":" & gstrServerName & " 306 :You have been marked as being away"
        End If
      Case "VERSION":
        ' RPL_VERSION
        SendData ":" & gstrServerName & " 351 " & App.Title & "|" & gstrServerVersion & " " & gstrServerName & " :" & gstrServerComments
      Case "STATS":
        Select Case LCase(aryCmds(1))
          Case "":
            ' ERR_NEEDMOREPARAMS
            SendData ":" & gstrServerName & " 461 STATS :Not enough parameters"
          Case "u":
            Dim lngDays     As Long
            Dim lngHours    As Long
            Dim lngMinutes  As Long
            Dim lngSeconds  As Long
          
            ' RPL_STATSUPTIME
            SendData ":" & gstrServerName & " 242 :Server Up " & lngDays & " days " & Format(lngHours, "00") & ":" & Format(lngMinutes, "00") & ":" & Format(lngSeconds, "00")
            ' RPL_ENDOFSTATS
            SendData ":" & gstrServerName & " 219 u :End of STATS report"
        End Select
      Case "PRIVMSG", "NOTICE":
        Select Case True
          Case (Len(aryCmds(1)) = 0) And (strCmd <> "NOTICE"):
            ' ERR_NEEDMOREPARAMS
            SendData ":" & gstrServerName & " 461 PRIVMSG :Not enough parameters"
          Case (Len(aryCmds(2)) = 0) And (strCmd <> "NOTICE"):
            ' ERR_NOTEXTTOSEND
            SendData ":" & gstrServerName & " 412 :No text to send"
          Case IsConnect(aryCmds(1)):
            SendConnectPriv Me.FullIdent, aryCmds(1), aryCmds(2), , strCmd
          Case IsChannel(aryCmds(1)):
            SendChannelPriv Me.FullIdent, aryCmds(1), aryCmds(2), , strCmd
          Case (strCmd <> "NOTICE"):
            ' ERR_NOSUCHNICK
            SendData ":" & gstrServerName & " 401 " & aryCmds(1) & " :No such nick/channel"
        End Select
      Case "WHOIS":
        Whois strNick, aryCmds(1)
      Case Else:
        ' ERR_UNKNOWNCOMMAND
        SendData ":" & gstrServerName & " 421 " & UCase(strCmd) & " :Unknown command (" & UCase(strCmd) & ")"
    End Select
    
NextCmd:
    
    If (strCmd <> "PING") And (strCmd <> "PONG") Then
      lngIdleTimer = Timer
    End If
  Next lngCmdLnIndex
End Sub

Private Sub ChangeNick(strNewNick As String)
  Select Case True
    Case (Len(strNewNick) = 0):
      ' ERR_NONICKNAMEGIVEN
      SendData ":" & gstrServerName & " 431 :No nickname given"
    Case (Not IsValidNick(strNewNick)):
      ' ERR_ERRONEUSNICKNAME
      SendData ":" & gstrServerName & " 432 " & strNewNick & " :Erroneus nickname"
    Case NickInUse(strNewNick):
      ' ERR_NICKNAMEINUSE
      SendData ":" & gstrServerName & " 433 " & strNewNick & " :Nickname is already in use"
    Case Else:
      If Len(strNick) = 0 Then
        strNick = strNewNick
      End If
    
      strNick = strNewNick
      
      If bolGotIdent And Not bolAuthed Then
        FinishAuth
      Else
        SendData ":" & strNick & " NICK " & strNewNick
      End If
      
      DebugPrint "[" & strDebugGUID & " |NICK  ] " & strNick
  End Select
End Sub

Private Sub SetUserInfo(strUser As String, strHost As String, strServer As String, strRName As String)
  If bolAuthed Then
    ' ERR_ALREADYREGISTRED
    SendData ":" & gstrServerName & " 462 :You may not reregister"
  End If

  Dim bolNotEnoughParam As Boolean
  bolNotEnoughParam = True

  Select Case 0
    Case Len(strUser):
    Case Len(strHost):
    Case Len(strServer):
    Case Len(strRName):
    Case Else
      bolNotEnoughParam = False
  End Select
  
  If bolNotEnoughParam Then
    ' ERR_NEEDMOREPARAMS
    SendData ":" & gstrServerName & " 461 USER :Not enough parameters"
    
    Exit Sub
  End If
  
  strUsername = strUser
  
  strHostName = AddressToName(sckConnect.RemoteHostIP)
  If Len(strHostName) = 0 Then
    strHostName = sckConnect.RemoteHost
  End If
  
  strRealName = strRName
  
  bolGotIdent = True
  
  If bolGotIdent And (Not bolAuthed) And (Len(strNick) > 0) Then
    FinishAuth
  End If
  
  DebugPrint "[" & strDebugGUID & " |USER  ] User Information Recieved"
End Sub

Public Sub Quit(strMsg As String)
  On Error Resume Next

  sckConnect.CloseSocket
  
  SendNickChannel strNick, ":" & Me.FullIdent & " QUIT :" & strMsg
  RemoveNickChannel strNick
  
  DestroyConnect strNick
End Sub

Private Sub FinishAuth()
  SendWelcome strNick
  SendMOTD strNick
  
  PingPong strNick
          
  SendData ":" & strNick & " MODE " & strNick & " +" & strModes

  bolAuthed = True
End Sub

Private Sub tmrConnect_Timer()
  On Error GoTo ErrH

  If Not bolPong Then
    Quit "Ping timed out"
  End If

  PingPong strNick
  
  bolPong = False
  
  If glngPingInterval <> tmrConnect.Interval Then
    tmrConnect.Interval = glngPingInterval
  End If
  
  Exit Sub
ErrH:
  Quit Err.Description
  
  Exit Sub
End Sub

Private Sub tmrSend_Timer()
  Dim strData             As String
  strData = stkSend.NextItem

  If Len(strData) > 0 Then
    sckConnect.SendData strData
  
    If gbolShowSentData Then
      DebugPrint "[" & strDebugGUID & " |SENT  ] " & strData
    End If
  End If
  
  If glngSendInterval <> tmrSend.Interval Then
    tmrSend.Interval = glngSendInterval
  End If
End Sub
