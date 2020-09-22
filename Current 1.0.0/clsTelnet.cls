VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum enumAction
  Username
  Password
  Command
  Wait
End Enum

Private WithEvents sckTelnet  As clsSocket
Attribute sckTelnet.VB_VarHelpID = -1
Private WithEvents sckConnect As clsSocket
Attribute sckConnect.VB_VarHelpID = -1

Private strGUID               As String
Private strDebugGUID          As String

Private actCurrent            As enumAction

Private strUsername           As String
Private strTelnetIdent        As String

Private bolConsoleMode        As Boolean

Public Property Get GUID() As String
  GUID = strGUID
End Property

Public Property Get IP() As String
  IP = sckTelnet.LocalIP
End Property

Public Property Get Port() As Long
  Port = sckTelnet.LocalPort
End Property

Public Property Get ConsoleMode() As Boolean
  ConsoleMode = bolConsoleMode
End Property

Public Function Initialize(lngPort As Long, Optional strIP As String)
  On Error GoTo ErrH

  strGUID = GenerateGUID

  Dim strTelnet As String
  strTelnet = Space(20)
  
  strDebugGUID = Replace(strGUID, "-", "")
  strDebugGUID = Replace(strDebugGUID, "}", "")
  strDebugGUID = Replace(strDebugGUID, "{", "")
  strDebugGUID = Left(strDebugGUID, (Len(strDebugGUID) / 2) - 6)
  strDebugGUID = strTelnet & "|" & strDebugGUID
  
  Set sckConnect = New clsSocket
  sckConnect.Protocol = sckTCPProtocol
  
  Set sckTelnet = New clsSocket
  sckTelnet.Protocol = sckTCPProtocol
  
  If Len(strIP) = 0 Then
    strIP = sckTelnet.LocalIP
  End If
  
  sckTelnet.Bind lngPort, strIP
  sckTelnet.Listen
  
  LSet strTelnet = "TELN " & sckTelnet.LocalIP

  DebugPrint "[" & strDebugGUID & " |INIT  ] Telnet Administration"
  DebugPrint "[" & strDebugGUID & " |INIT  ]   GUID:      " & strGUID
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Hostname:  " & sckTelnet.LocalHostName
  DebugPrint "[" & strDebugGUID & " |INIT  ]   IP:        " & sckTelnet.LocalIP
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Port:      " & sckTelnet.LocalPort
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Thread ID: " & GetCurrentThreadId
  
  Exit Function
ErrH:
  Dim errCurrent As ErrObject
  Set errCurrent = Err

  DebugPrint "[" & strDebugGUID & " |ERROR ] Telnet Initialization Error"
  DebugPrint "[" & strDebugGUID & " |ERROR ]   Number:      " & errCurrent.Number
  DebugPrint "[" & strDebugGUID & " |ERROR ]   Description: " & errCurrent.Description
  DebugPrint "[" & strDebugGUID & " |ERROR ]   Source:      " & errCurrent.Source
  
  Exit Function
End Function

Public Function SendData(strData As String)
  If sckConnect.State = sckConnected Then
    sckConnect.SendData strData
  End If
End Function

Private Sub Class_Terminate()
  If sckTelnet.State = sckConnected Then
    sckTelnet.SendData vbCrLf
    sckTelnet.SendData "Server is shutting down." & vbCrLf
  End If

  sckTelnet.CloseSocket
  Set sckTelnet = Nothing

  DebugPrint "[" & strDebugGUID & " |CLOSE ] Terminated"
End Sub

Private Sub sckConnect_OnClose()
  DebugPrint "[" & strDebugGUID & " |CLOSE ] Telnet Session Terminated."
End Sub

Private Sub sckConnect_OnDataArrival(ByVal bytesTotal As Long)
  On Error Resume Next

  Dim strData     As String
  sckConnect.GetData strData, vbString
  
  Static strLine  As String
  
  Select Case strData
    Case Chr(8):
      strLine = Left(strLine, Len(strLine) - 1)
    Case vbCrLf:
      DoCommand strLine, sckConnect
      
      strLine = ""
    Case Else:
      strLine = strLine & strData
  End Select
End Sub

Private Sub sckTelnet_OnConnectionRequest(ByVal requestID As Long)
  DebugPrint "[" & strDebugGUID & " |CONREQ] RequestID: " & requestID
  
  If sckConnect.State = sckConnected Then
    DebugPrint "[" & strDebugGUID & " |DENY  ] Too many telnet admin connections."
    
    Dim sckDeny As clsSocket
    
    Set sckDeny = New clsSocket
    sckDeny.Protocol = sckTCPProtocol
    
    sckDeny.Accept requestID
    
    Do Until sckDeny.State = sckConnected
      DoEvents
    Loop
    
    sckDeny.SendData "Too many admin connections ..."
    
    Dim dblTimer As Double
    dblTimer = Timer + 2
    
    Do Until Timer > dblTimer
      DoEvents
    Loop
    
    sckDeny.CloseSocket
    Set sckDeny = Nothing
    
    Exit Sub
  Else
    DebugPrint "[" & strDebugGUID & " |TELNET] Telnet Administrator Connected."
    
    bolConsoleMode = False
    
    strUsername = ""
    strTelnetIdent = ""
  
    sckConnect.CloseSocket
    sckConnect.Accept requestID
    
    Do Until sckConnect.State = sckConnected
      DoEvents
    Loop
    
    TelnetWelcome sckConnect
  End If
End Sub

Private Function DoCommand(strCommand As String, sckSend As clsSocket)
  Select Case actCurrent
    Case Username:
      strUsername = Replace(strCommand, " ", "")
      
      sckConnect.SendData "Password: "
        
      actCurrent = Password
    
      DebugPrint "[" & strDebugGUID & " |USER  ] Telnet Username recieved:  " & strUsername
    Case Password:
      Static lngTries As Long
      
      TelnetCLS sckSend
      
      Dim logUser As clsLogin
      Set logUser = New clsLogin
    
      If logUser.VerifyLogin(strUsername, "", strCommand) Then
        
        DebugPrint "[" & strDebugGUID & " |PASS  ] Telnet Password recieved. (valid)"
        
        sckConnect.SendData "Authentication Successful!" & vbCrLf
        sckConnect.SendData vbCrLf
        
        actCurrent = Command
      
        strTelnetIdent = "TELNET|" & strUsername & "!" & sckConnect.RemoteHostIP & "@" & sckConnect.RemoteHost
        
        sckSend.SendData gstrServerName & "> "
      
        DebugPrint "[" & strDebugGUID & " |TELNET] Telnet Command Prompt Activated."
      Else
        DebugPrint "[" & strDebugGUID & " |PASS  ] Telnet Password recieved. (invalid)"
        
        strUsername = ""
        strTelnetIdent = ""
      
        sckConnect.SendData "Invalid Username/Password!" & vbCrLf
        sckConnect.SendData vbCrLf
        
        If lngTries = 2 Then
          sckConnect.CloseSocket
          
          lngTries = 0
        Else
          sckConnect.SendData "Username: "
          
          lngTries = lngTries + 1
        
          actCurrent = Username
        End If
      End If
      
      Set logUser = Nothing
    Case Wait:
      actCurrent = Command
    Case Command:
      DebugPrint "[" & strDebugGUID & " |COMMND] Telnet Command recieved:  " & strCommand
      
      Dim strCmd    As String
      If InStr(strCommand, " ") > 0 Then
        strCmd = Left(strCommand, InStr(strCommand, " ") - 1)
      Else
        strCmd = strCommand
      End If
      
      strCommand = Trim(Right(strCommand, Len(strCommand) - Len(strCmd)))
      
      Dim aryCmd()  As String
      aryCmd = Split(strCommand, " ")
      
      ReDim Preserve aryCmd(0 To 10) As String
    
      Select Case UCase(strCmd)
        Case "CHANNELS":
          sckSend.SendData " -===============================================-" & vbCrLf
          sckSend.SendData " Active Channels:  " & Format(colChannels.Count, "#,##0") & vbCrLf
          
          Dim chnCheck      As clsChannel
          
          Dim lngChannelLineCount  As Long
          lngChannelLineCount = 1
          
          For Each chnCheck In colChannels
            sckSend.SendData "   " & chnCheck.Name & vbCrLf
            
            If lngChannelLineCount = 23 Then
              sckSend.SendData vbTab & vbTab & "-Press Any Key-" & vbCrLf
              
              actCurrent = Wait
              
              Do Until actCurrent = Command
                DoEvents
              Loop
              
              lngChannelLineCount = 1
            Else
              lngChannelLineCount = lngChannelLineCount + 1
            End If
          Next
          sckSend.SendData " -===============================================-" & vbCrLf
        Case "CONNECTS":
          sckSend.SendData " -===============================================-" & vbCrLf
          sckSend.SendData " Active Connections:  " & Format(colConnects.Count, "#,##0") & vbCrLf
          
          Dim conCheck      As clsConnect
          
          Dim lngConnectLineCount  As Long
          lngConnectLineCount = 1
          
          For Each conCheck In colConnects
            sckSend.SendData "   " & conCheck.FullIdent & vbCrLf
            
            If lngConnectLineCount = 23 Then
              sckSend.SendData vbTab & vbTab & "-Press Any Key-" & vbCrLf
              
              actCurrent = Wait
              
              Do Until actCurrent = Command
                DoEvents
              Loop
              
              lngConnectLineCount = 1
            Else
              lngConnectLineCount = lngConnectLineCount + 1
            End If
          Next
          sckSend.SendData " -===============================================-" & vbCrLf
        Case "STATS":
          sckSend.SendData " -===============================================-" & vbCrLf
          sckSend.SendData " Server Statistics:" & vbCrLf
          sckSend.SendData "    Channels:         " & Format(colChannels.Count, "#,##0") & vbCrLf
          sckSend.SendData "    Connections:      " & Format(colConnects.Count, "#,##0") & vbCrLf
          sckSend.SendData vbCrLf
          sckSend.SendData "    Name:             " & gstrServerName & vbCrLf
          sckSend.SendData "    Version:          " & gstrServerVersion & vbCrLf
          sckSend.SendData "    StartTime:        " & gstrServerStartTime & vbCrLf
          sckSend.SendData "    Comments:         " & gstrServerComments & vbCrLf
          sckSend.SendData "    Network Name:     " & gstrNetworkName & vbCrLf
          sckSend.SendData vbCrLf
          sckSend.SendData "    Ping Interval:    " & Format(glngPingInterval, "#,##0") & " ms" & vbCrLf
          sckSend.SendData "    Send Interval:    " & Format(glngSendInterval, "#,##0") & " ms" & vbCrLf
          sckSend.SendData vbCrLf
          sckSend.SendData "    Max Connections:  " & Format(glngMaxConnections, "#,##0") & vbCrLf
          sckSend.SendData "    Max Joins:        " & Format(glngMaxJoins, "#,##0") & vbCrLf
          sckSend.SendData vbCrLf
          sckSend.SendData " IRC Server IP Address(s):" & vbCrLf
          
          Dim svrCheck As clsListen
          
          For Each svrCheck In colServers
            sckSend.SendData "   " & svrCheck.IP & ":" & svrCheck.Port & vbCrLf
          Next
          
          sckSend.SendData vbCrLf
          sckSend.SendData " Telnet Server IP Address(s):" & vbCrLf
          
          Dim telCheck As clsTelnet
          
          For Each telCheck In colTelnet
            sckSend.SendData "   " & telCheck.IP & ":" & telCheck.Port & vbCrLf
          Next
          
          sckSend.SendData " -===============================================-" & vbCrLf
        Case "KILL", "DISCONNECT"
          If (Len(aryCmd(0)) > 0) Then
            If IsConnect(aryCmd(0)) Then
              SendConnectPriv strTelnetIdent, aryCmd(0), "You are being disconnected by this Telnet Operator.", False
              
              Dim conTarget As clsConnect
              Set conTarget = GetConnect(aryCmd(0))
              
              conTarget.Quit "Disconnected by Telnet Operator: " & strUsername
            
              Set conTarget = Nothing
              
              sckSend.SendData aryCmd(0) & " was disconnected from this server." & vbCrLf
            Else
              sckSend.SendData "Invalid Target." & vbCrLf
            End If
          Else
            sckSend.SendData "Insufficient Parameters." & vbCrLf
          End If
        Case "KICK":
          If (Len(aryCmd(0)) > 0) And (Len(aryCmd(1)) > 0) Then
            If IsConnect(aryCmd(0)) Then
              KickNick strTelnetIdent, aryCmd(0), aryCmd(1), "Telnet Operator has kicked you from " & aryCmd(0) & "."
                
              sckSend.SendData aryCmd(1) & " was kicked from " & aryCmd(0) & "." & vbCrLf
            Else
              sckSend.SendData "Invalid Target." & vbCrLf
            End If
          Else
            sckSend.SendData "Insufficient Parameters." & vbCrLf
          End If
        Case "MSG", "PRIVMSG":
          If (Len(aryCmd(0)) > 0) And (Len(aryCmd(1)) > 0) Then
            If IsConnect(aryCmd(0)) Then
              If SendConnectPriv(strTelnetIdent, aryCmd(0), aryCmd(1), False) Then
                sckSend.SendData "Message sent successfully." & vbCrLf
              Else
                sckSend.SendData "Could not send message." & vbCrLf
              End If
            ElseIf IsChannel(aryCmd(0)) Then
              If SendChannelPriv("TELNET|" & strUsername & "!" & sckConnect.RemoteHostIP & "@" & sckConnect.RemoteHost, aryCmd(0), aryCmd(1), False) Then
                sckSend.SendData "Message sent successfully." & vbCrLf
              Else
                sckSend.SendData "Could not send message." & vbCrLf
              End If
            Else
              sckSend.SendData "Invalid Target." & vbCrLf
            End If
          Else
            sckSend.SendData "Insufficient Parameters." & vbCrLf
          End If
        Case "CONSOLEMODE":
          bolConsoleMode = Not bolConsoleMode
          
          If bolConsoleMode Then
            sckSend.SendData " Console Mode Activated - console output will be sent" & vbCrLf
          Else
            sckSend.SendData vbCrLf
            sckSend.SendData " Console Mode Deactivated - console output will not be sent" & vbCrLf
          End If
        Case "CLS", "CLEARSCREEN":
          TelnetCLS sckConnect
        Case "INFO", "ABOUT":
          sckSend.SendData " -===============================================-" & vbCrLf
          sckSend.SendData " -=                  Telnet Server              =-" & vbCrLf
          sckSend.SendData " -=              Version: " & gstrServerVersion & "              =-" & vbCrLf
          sckSend.SendData " -===============================================-" & vbCrLf
        Case "QUIT", "EXIT":
          sckSend.SendData "Disconnecting ..." & vbCrLf
          sckSend.CloseSocket
        Case "HELP", "?":
          sckSend.SendData " -=============================================================-" & vbCrLf
          sckSend.SendData " -= CHANNELS - Displays list of all active channels currently =-" & vbCrLf
          sckSend.SendData " -=             in use on this IRC server.                    =-" & vbCrLf
          sckSend.SendData " -= CONNECTS - Displays list of all active connections        =-" & vbCrLf
          sckSend.SendData " -=             to this IRC server.                           =-" & vbCrLf
          sckSend.SendData " -= STATS    - Displays statistics on the current IRC server. =-" & vbCrLf
          sckSend.SendData " -= KILL <nickname>     or     DISCONNECT <nickname>          =-" & vbCrLf
          sckSend.SendData " -=          - Disconnects the target nickname from this      =-" & vbCrLf
          sckSend.SendData " -=             IRC server.                                   =-" & vbCrLf
          sckSend.SendData " -= KICK <channel> <nickname>                                 =-" & vbCrLf
          sckSend.SendData " -=          - Kicks the target nickname from the target      =-" & vbCrLf
          sckSend.SendData " -=             channel.                                      =-" & vbCrLf
          sckSend.SendData " -= MSG <nickname> <msg>   or   PRIVMSG <nickname> <msg>      =-" & vbCrLf
          sckSend.SendData " -=          - Sends a private message to the target nickname =-" & vbCrLf
          sckSend.SendData " -=             channels can also be used as targets.         =-" & vbCrLf
          sckSend.SendData " -= CONSOLEMODE                                               =-" & vbCrLf
          sckSend.SendData " -=          - Sets the telnet session into console mode, all =-" & vbCrLf
          sckSend.SendData " -=             server console output is forwarded to the     =-" & vbCrLf
          sckSend.SendData " -=             telnet console. This command toggles on/off.  =-" & vbCrLf
          sckSend.SendData " -= CLS      - Clears the telnet session's display.           =-" & vbCrLf
          sckSend.SendData " -= INFO     - Displays version information.                  =-" & vbCrLf
          sckSend.SendData " -= HELP2    - Displays more help information.                =-" & vbCrLf
          sckSend.SendData " -=============================================================-" & vbCrLf
        Case "HELP2", "?2":
          sckSend.SendData " -=============================================================-" & vbCrLf
          sckSend.SendData " -= SERVERNAME <name>                                         =-" & vbCrLf
          sckSend.SendData " -=            Changes the current name of the server.        =-" & vbCrLf
          sckSend.SendData " -= SERVERCOMMENTS <comments>                                 =-" & vbCrLf
          sckSend.SendData " -=            Changes the current comments of the server.    =-" & vbCrLf
          sckSend.SendData " -= NETWORKNAME <name>                                        =-" & vbCrLf
          sckSend.SendData " -=            Changes the name of the network that the       =-" & vbCrLf
          sckSend.SendData " -=             server belongs to.                            =-" & vbCrLf
          sckSend.SendData " -= PINGINTERVAL <interval>                                   =-" & vbCrLf
          sckSend.SendData " -=            Changes the interval at which clients are      =-" & vbCrLf
          sckSend.SendData " -=             pinged for inactivity and dropped connects.   =-" & vbCrLf
          sckSend.SendData " -= SENDINTERVAL <interval>                                   =-" & vbCrLf
          sckSend.SendData " -=            Changes the interval at which data packets     =-" & vbCrLf
          sckSend.SendData " -=             are sent to connected clients.                =-" & vbCrLf
          sckSend.SendData " -= MAXCONNECTIONS <max>                                      =-" & vbCrLf
          sckSend.SendData " -=            Changes the maximum connections allowed.       =-" & vbCrLf
          sckSend.SendData " -= MAXJOINS <max>                                            =-" & vbCrLf
          sckSend.SendData " -=            Changes the maximum number of channels any one =-" & vbCrLf
          sckSend.SendData " -=             client is allowed to be a member of.          =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -= HELP3    - Displays more helpl information.               =-" & vbCrLf
          sckSend.SendData " -=============================================================-" & vbCrLf
        Case "HELP3", "?3":
          sckSend.SendData " -=============================================================-" & vbCrLf
          sckSend.SendData " -= SHOWSENTDATA <1/0>                                        =-" & vbCrLf
          sckSend.SendData " -=            Show all sent data in server console.          =-" & vbCrLf
          sckSend.SendData " -= SHOWRCVDDATA <1/0>                                        =-" & vbCrLf
          sckSend.SendData " -=            Show all recieved data in server console.      =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -=                                                           =-" & vbCrLf
          sckSend.SendData " -= SHUTDOWN - Shuts down IRC server. Also terminates telnet  =-" & vbCrLf
          sckSend.SendData " -=             sessions.                                     =-" & vbCrLf
          sckSend.SendData " -= QUIT     - Terminates the current telnet session.         =-" & vbCrLf
          sckSend.SendData " -=============================================================-" & vbCrLf
        Case "SERVERNAME":
          If Len(strCommand) > 0 Then
            gstrServerName = Replace(strCommand, " ", "_")
          End If
            
          sckSend.SendData "  " & gstrServerName & vbCrLf
        Case "SERVERCOMMENTS":
          If Len(strCommand) > 0 Then
            gstrServerComments = strCommand
          End If
            
          sckSend.SendData "  " & gstrServerComments & vbCrLf
        Case "SERVERCOMMENTS":
          If Len(strCommand) > 0 Then
            gstrNetworkName = Replace(strCommand, " ", "_")
          End If
            
          sckSend.SendData "  " & gstrNetworkName & vbCrLf
        Case "PINGINTERVAL":
          If Len(aryCmd(0)) > 0 Then
            glngPingInterval = Val(aryCmd(0))
          End If
            
          sckSend.SendData "  " & glngPingInterval & vbCrLf
        Case "SENDINTERVAL":
          If Len(aryCmd(0)) > 0 Then
            glngSendInterval = Val(aryCmd(0))
          End If
            
          sckSend.SendData "  " & glngSendInterval & vbCrLf
        Case "MAXCONNECTIONS":
          If Len(aryCmd(0)) > 0 Then
            glngMaxConnections = Val(aryCmd(0))
          End If
            
          sckSend.SendData "  " & glngMaxConnections & vbCrLf
        Case "MAXJOINS":
          If Len(aryCmd(0)) > 0 Then
            glngMaxJoins = Val(aryCmd(0))
          End If
            
          sckSend.SendData "  " & glngMaxJoins & vbCrLf
        Case "SHOWSENTDATA":
          If Len(aryCmd(0)) > 0 Then
            gbolShowSentData = Val(aryCmd(0))
          End If
        
          sckSend.SendData "  " & gbolShowSentData & vbCrLf
        Case "SHOWRCVDDATA":
          If Len(aryCmd(0)) > 0 Then
            gbolShowRcvdData = Val(aryCmd(0))
          End If
        
          sckSend.SendData "  " & gbolShowRcvdData & vbCrLf
        Case "SHUTDOWN":
          Unload frmMain
        Case "":
        Case Else:
          sckSend.SendData " Invalid Command." & vbCrLf & vbCrLf
      End Select
      
      If (sckSend.State = sckConnected) And Not bolConsoleMode Then
        sckSend.SendData gstrServerName & "> "
      End If
  End Select
End Function

Private Function TelnetWelcome(sckSend As clsSocket)
  TelnetCLS sckSend
  
  Dim strTitle      As String
  strTitle = gstrServerName & "| Telnet Administration"
  
  Dim lngSideBuffer As Long
  lngSideBuffer = ((45 - Len(strTitle)) / 2)
  
  strTitle = Space(lngSideBuffer) & strTitle & Space(lngSideBuffer)
  
  sckSend.SendData vbTab & vbTab & strTitle & vbCrLf
  sckSend.SendData vbTab & vbTab & "    *    Authorized connections only    *    " & vbCrLf
  sckSend.SendData vbTab & vbTab & "    *     User Access Verification      *    " & vbCrLf
  
  sckSend.SendData vbCrLf
  sckSend.SendData "Username: "
  
  actCurrent = Username

  DebugPrint "[" & strDebugGUID & " |HELLO ] Telnet Welcome and Login Prompt sent."
End Function

Private Function TelnetCLS(sckSend As clsSocket)
  sckSend.SendData Chr$(27) & "[2J"
  
  DoEvents
End Function
