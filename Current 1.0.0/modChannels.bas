Attribute VB_Name = "modChannels"
Option Explicit

' Channel Replies, may or may not be implemented
'           ERR_NEEDMOREPARAMS  ERR_BANNEDFROMCHAN
'           ERR_INVITEONLYCHAN  ERR_BADCHANNELKEY
'           ERR_CHANNELISFULL   ERR_BADCHANMASK
'           ERR_NOSUCHCHANNEL   ERR_TOOMANYCHANNELS

'405             ERR_TOOMANYCHANNELS
'                        "<channel name> :You have joined too many \
'                         channels ""


Public Function CreateChannel(strName As String) As String
  Dim chnNew As clsChannel
  Set chnNew = New clsChannel
  
  chnNew.Initialize strName
  
  colChannels.Add chnNew, chnNew.GUID
End Function

Public Function GetChannel(strName As String) As clsChannel
  Dim chnCheck As clsChannel
  
  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colChannels.Count
    Set chnCheck = colChannels(lngIndex)
    
    If LCase(chnCheck.Name) = LCase(strName) Then
      Set GetChannel = chnCheck
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Public Function DestroyChannel(strName As String)
  Dim chnCheck As clsChannel
  
  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colChannels.Count
    Set chnCheck = colChannels(lngIndex)
    
    If LCase(chnCheck.Name) = LCase(strName) Then
      colChannels.Remove lngIndex
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Public Function JoinChannel(strName As String, strNick As String)
  If Left(strName, 1) <> "#" Then
    Dim conJoin As clsConnect
    Set conJoin = GetConnect(strNick)
    
    ' ERR_ERRONEUSNICKNAME
    conJoin.SendData ":" & gstrServerName & " 432 " & strName & " :Erroneus channel name"
    
    Exit Function
  End If

  Dim chnJoin As clsChannel
  Set chnJoin = GetChannel(strName)
  
  If chnJoin Is Nothing Then
    CreateChannel strName
    
    Set chnJoin = GetChannel(strName)
  End If
  
  chnJoin.JoinNick strNick
End Function

Public Function PartChannel(strName As String, strNick As String, strMsg As String)
  Dim chnPart As clsChannel
  Set chnPart = GetChannel(strName)
  
  If chnPart Is Nothing Then
    Dim conNick As clsConnect
    Set conNick = GetConnect(strNick)
    
    ' ERR_NOSUCHCHANNEL
    conNick.SendData ":" & gstrServerName & " 403 " & strName & " :No such channel"
  Else
    chnPart.PartNick strNick, strMsg
  End If
End Function

Public Function KickNick(strByNick As String, strName As String, strNick As String, strMsg As String)
  Dim chnKick As clsChannel
  Set chnKick = GetChannel(strName)
  
  If chnKick Is Nothing Then
    Dim conNick As clsConnect
    Set conNick = GetConnect(strByNick)
    
    ' ERR_NOSUCHCHANNEL
    conNick.SendData ":" & gstrServerName & " 403 " & strName & " :No such channel"
  Else
    If InStr(chnKick.NickMode(strByNick), "o") = 0 Then
      ' ERR_CHANOPRIVSNEEDED
      conNick.SendData ":" & gstrServerName & " 482 " & chnKick.Name & " :You're not channel operator"
    Else
      chnKick.KickNick strByNick, strNick, strMsg
    End If
  End If
End Function

Public Function SetChanTopic(strName As String, strNick As String, strTopic As String)
  Dim chnTopic As clsChannel
  Set chnTopic = GetChannel(strName)
  
  If chnTopic Is Nothing Then
    Dim conNick As clsConnect
    Set conNick = GetConnect(strNick)
    
    ' ERR_NOSUCHCHANNEL
    conNick.SendData ":" & gstrServerName & " 403 " & strName & " :No such channel"
  Else
    chnTopic.SetTopic strNick, strTopic
  End If
End Function

Public Function IsChannel(strName As String) As Boolean
  IsChannel = Not (GetChannel(strName) Is Nothing)
End Function

Public Function SendChannelPriv(strFrom As String, strName As String, strMsg As String, Optional bolSendErrors As Boolean = True, Optional strCmd As String = "PRIVMSG")
  Dim conFrom As clsConnect
  Set conFrom = GetConnect(strFrom)

  Dim chnSend As clsChannel
  Set chnSend = GetChannel(strName)

  If (InStr(chnSend.Mode, "n") > 0) And Not chnSend.HaveNick(strFrom) Then
    ' ERR_CANNOTSENDTOCHAN
    conFrom.SendData ":" & gstrServerName & " 404 " & chnSend.Name & " :Cannot send to channel (+n)"
  ElseIf (InStr(chnSend.Mode, "m") > 0) And Not ((InStr(chnSend.NickMode(strFrom), "o") > 0) Or (InStr(chnSend.NickMode(strFrom), "v") > 0)) Then
    ' ERR_CANNOTSENDTOCHAN
    conFrom.SendData ":" & gstrServerName & " 404 " & chnSend.Name & " :Cannot send to channel (+m)"
  Else
    chnSend.SendData ":" & conFrom.FullIdent & " " & strCmd & " " & chnSend.Name & " :" & strMsg
  End If
End Function

Public Function SendNickChannel(strNick As String, strData As String)
  Dim chnCheck As clsChannel
  
  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colChannels.Count
    Set chnCheck = colChannels(lngIndex)
    
    If chnCheck.HaveNick(strNick) Then
      chnCheck.SendData strData
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Public Function RemoveNickChannel(strNick As String)
  Dim chnCheck As clsChannel
  
  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colChannels.Count
    Set chnCheck = colChannels(lngIndex)
    
    If chnCheck.HaveNick(strNick) Then
      chnCheck.RemoveNick strNick
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

