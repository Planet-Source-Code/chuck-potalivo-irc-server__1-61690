Attribute VB_Name = "modConnects"
Option Explicit

Public Function CreateConnect(lngRequestID As Long)
  Dim bolConnectOk As Boolean
  bolConnectOk = (colConnects.Count < glngMaxConnections)

  Dim conNew As clsConnect
  Set conNew = New clsConnect
  
  conNew.Initialize lngRequestID, bolConnectOk
  
  If bolConnectOk Then
    colConnects.Add conNew, conNew.GUID
  Else
    ' RPL_BOUNCE
    conNew.SendData ":" & gstrServerName & " 005 ""To many connections, please try again later."""
    
    Dim dblTime As Double
    dblTime = Timer + 5
    
    Do Until Timer > dblTime
      DoEvents
    Loop
    
    Set conNew = Nothing
  End If
End Function

Public Function DestroyConnect(strNick As String)
  Dim conCheck As clsConnect
  
  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colConnects.Count
    Set conCheck = colConnects(lngIndex)
    
    If LCase(conCheck.Nick) = LCase(strNick) Then
      colConnects.Remove lngIndex
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Public Function NickInUse(strNick As String) As Boolean
  Dim conCheck As clsConnect
  Set conCheck = GetConnect(strNick)
  
  NickInUse = Not (conCheck Is Nothing)
End Function

Public Function GetConnect(strNick As String) As clsConnect
  Dim conCheck As clsConnect

  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colConnects.Count
    Set conCheck = colConnects(lngIndex)
    
    If (LCase(conCheck.Nick) = LCase(strNick)) Or (LCase(conCheck.FullIdent) = LCase(strNick)) Then
      Set GetConnect = conCheck
      
      Exit Do
    End If
  
    lngIndex = lngIndex + 1
  Loop
End Function

Public Function PingPong(strNick As String)
  Dim conNick As clsConnect
  Set conNick = GetConnect(strNick)
  
  If conNick Is Nothing Then
  Else
    conNick.SendData "PING " & Round(Timer, 6)
  End If
End Function

Public Function IsConnect(strNick As String) As Boolean
  IsConnect = NickInUse(strNick)
End Function

Public Function SendConnectPriv(strFrom As String, strNick As String, strMsg As String, Optional bolSendErrors As Boolean = True, Optional strCmd As String = "PRIVMSG") As Boolean
  ' The following are reply messages that must be integrated still
  '
  '407         ERR_TOOMANYTARGETS
  '            "<target> :<error code> recipients. <abort message>"
  '
  '       - Returned to a client which is attempting to send a
  '         PRIVMSG/NOTICE using the user@host destination format
  '         and for a user@host which has several occurrences.
  '
  '       - Returned to a client which trying to send a
  '         PRIVMSG/NOTICE to too many recipients.
  '
  '       - Returned to a client which is attempting to JOIN a safe
  '         channel using the shortname when there are more than one
  '         such channel.
  '
  '411         ERR_NORECIPIENT
  '            ":No recipient given (<command>)"
  '
  '412         ERR_NOTEXTTOSEND
  '            ":No text to send"
  '
  '415         ERR_BADMASK
  '            "<mask> :Bad Server/host mask"
  '
  '      - 412 - 415 are returned by PRIVMSG to indicate that
  '         the message wasn't delivered for some reason.
  '         ERR_NOTOPLEVEL and ERR_WILDTOPLEVEL are errors that
  '         are returned when an invalid use of
  '         "PRIVMSG $<server>" or "PRIVMSG #<host>" is attempted.


  Dim conFrom As clsConnect
  Set conFrom = GetConnect(strFrom)

  Dim conSend As clsConnect
  Set conSend = GetConnect(strNick)

  If Len(conSend.Away) > 0 Then
    If Not conFrom Is Nothing Then
      ' RPL_AWAY
      conFrom.SendData ":" & gstrServerName & " 301 " & conFrom.Nick & " " & conSend.Nick & " :" & conSend.Away
    End If
    
    SendConnectPriv = False
  Else
    Dim strFullIdent As String
    
    If conFrom Is Nothing Then
      strFullIdent = strFrom
    Else
      strFullIdent = conFrom.FullIdent
    End If
  
    conSend.SendData ":" & strFullIdent & " " & strCmd & " " & conSend.Nick & " :" & strMsg
    
    SendConnectPriv = True
  End If
End Function

Public Function Whois(strByNick As String, strNick As String)
  Dim conSend As clsConnect
  Set conSend = GetConnect(strByNick)
  
  Dim conWhois As clsConnect
  Set conWhois = GetConnect(strNick)
  
  If conWhois Is Nothing Then
    ' ERR_NOSUCHNICK
    conSend.SendData ":" & gstrServerName & " 401 " & strNick & " :No such nick/channel"
    
    Exit Function
  ElseIf Len(conWhois.Away) > 0 Then
    ' RPL_AWAY
    conSend.SendData ":" & gstrServerName & " 301 " & conWhois.Nick & " :" & conWhois.Away
  End If
  
  ' RPL_WHOISUSER
  conSend.SendData ":" & gstrServerName & " 311 " & conSend.Nick & " " & conWhois.Nick & " " & conWhois.Username & " " & conWhois.hostname & " * :" & conWhois.RealName
  ' RPL_WHOISIDLE
  conSend.SendData ":" & gstrServerName & " 317 " & conSend.Nick & " " & conWhois.Nick & " " & conWhois.IdleTime & " :seconds idle"
  
  Dim strChannels As String
  
  Dim chnWhois As clsChannel
  
  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colChannels.Count
    Set chnWhois = colChannels(lngIndex)
    
    If chnWhois.HaveNick(strNick) Then
      Select Case True
        Case (InStr(chnWhois.NickMode(strNick), "o") > 0):
          strChannels = strChannels & "@" & chnWhois.Name
        Case (InStr(chnWhois.NickMode(strNick), "v") > 0):
          strChannels = strChannels & "+" & chnWhois.Name
        Case Else
          strChannels = strChannels & chnWhois.Name
      End Select
      strChannels = strChannels & " "
    End If
  
    lngIndex = lngIndex + 1
  Loop
  strChannels = Trim(strChannels)
  
  If Len(strChannels) > 0 Then
    ' RPL_WHOISCHANNELS
    conSend.SendData ":" & gstrServerName & " 319 " & conSend.Nick & " " & conWhois.Nick & " :" & strChannels
  End If
  
  ' RPL_ENDOFWHOIS
  conSend.SendData ":" & gstrServerName & " 318 " & conSend.Nick & " " & conWhois.Nick & " :End of WHOIS list"
End Function
