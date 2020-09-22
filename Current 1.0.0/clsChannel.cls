VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private strGUID                   As String
Private strDebugGUID              As String

Private strName                   As String

Private strTopic                  As String
Private strTopicSetBy             As String
Private lngTopicSetTime           As Long

Private strModes                  As String

Private colMembers                As Collection

Private colInvited                As Collection
Private colBanned                 As Collection

Public Property Get GUID() As String
  GUID = strGUID
End Property

Public Property Get Name() As String
  Name = strName
End Property

Public Property Get Mode() As String
  Mode = strModes
End Property

Public Property Let Moderated(strNick As String, bolValue As Boolean)
  If InStr(Me.NickMode(strNick), "o") = 0 Then
    Dim conNick As clsConnect
    Set conNick = GetConnect(strNick)
    
    ' ERR_CHANOPRIVSNEEDED
    conNick.SendData ":" & gstrServerName & " 482 " & strName & " :You're not channel operator"
    
    Exit Property
  End If

  If Not bolValue And (InStr(strModes, "m") > 0) Then
    strModes = Replace(strModes, "m", "")
    
    SendData ":" & strNick & " MODE " & strName & " -m"
  ElseIf bolValue And (InStr(strModes, "m") = 0) Then
    strModes = strModes & "m"
  
    SendData ":" & strNick & " MODE " & strName & " +m"
  End If
End Property

Public Function NickMode(strNick As String) As String
  Dim xmlNick As clsSimpleXML

  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colMembers.Count
    Set xmlNick = colMembers(lngIndex)
    
    If (LCase(xmlNick.Tag) = LCase(strNick)) Or (LCase(xmlNick.Value("FullIdent")) = LCase(strNick)) Then
      NickMode = xmlNick.Value("Modes")
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Public Function Initialize(strChannelName As String)
  Dim strChannel As String
  strChannel = Space(20)
  
  LSet strChannel = "CHAN " & strChannelName

  strGUID = GenerateGUID

  strDebugGUID = Replace(strGUID, "-", "")
  strDebugGUID = Replace(strDebugGUID, "}", "")
  strDebugGUID = Replace(strDebugGUID, "{", "")
  strDebugGUID = Left(strDebugGUID, (Len(strDebugGUID) / 2) - 6)
  strDebugGUID = strChannel & "|" & strDebugGUID
  
  Set colMembers = New Collection
  
  Set colInvited = New Collection
  Set colBanned = New Collection
  
  strName = strChannelName
  
  strTopic = "Unregistered Channel"
  strTopicSetBy = gstrServerName
  lngTopicSetTime = UnixTime
  
  strModes = "t"
  
  DebugPrint "[" & strDebugGUID & " |INIT  ] Channel Class"
  DebugPrint "[" & strDebugGUID & " |INIT  ]   GUID:      " & strGUID
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Name:      " & strName
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Thread ID: " & GetCurrentThreadId
End Function

Private Sub Class_Terminate()
  Dim xmlTerm As clsSimpleXML
  Dim conTerm As clsConnect

  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colMembers.Count
    Set xmlTerm = colMembers(lngIndex)
    
    PartNick xmlTerm.Tag, "Channel terminating"
    
    lngIndex = lngIndex + 1
  Loop

  Set colMembers = Nothing
  
  DebugPrint "[" & strDebugGUID & " |CLOSE ] Terminated"
End Sub

Public Function BanNick(strByNick As String, strNick As String)
  If Len(strNick) = 0 Then
    Exit Function
  End If
  
  If InStr(Me.NickMode(strByNick), "o") = 0 Then
    Dim conNick As clsConnect
    Set conNick = GetConnect(strByNick)
    
    ' ERR_CHANOPRIVSNEEDED
    conNick.SendData ":" & gstrServerName & " 482 " & strName & " :You're not channel operator"
    
    Exit Function
  End If
  
  SendData ":" & strByNick & " MODE " & strName & " +b " & strNick
  
  Dim xmlBan As clsSimpleXML
  Set xmlBan = New clsSimpleXML
  
  xmlBan.Tag = strNick
  
  colBanned.Add xmlBan, xmlBan.Tag
End Function

Public Function KickNick(strByNick As String, strNick As String, strMsg As String)
  Dim xmlNick As clsSimpleXML

  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colMembers.Count
    Set xmlNick = colMembers(lngIndex)
    
    If LCase(xmlNick.Tag) = LCase(strNick) Then
      ' ??
      SendData ":" & strByNick & " KICK " & strName & " " & strNick & " :" & strMsg
    
      colMembers.Remove lngIndex
      
      DebugPrint "[" & strDebugGUID & " |KICK  ] " & strNick
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Public Function RemoveNick(strNick As String)
  Dim xmlNick As clsSimpleXML

  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colMembers.Count
    Set xmlNick = colMembers(lngIndex)
    
    If LCase(xmlNick.Tag) = LCase(strNick) Then
      colMembers.Remove lngIndex
      
      DebugPrint "[" & strDebugGUID & " |REMOVE] " & strNick
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Public Function InviteNick(strByNick As String, strNick As String)
  Dim conNick As clsConnect
  Set conNick = GetConnect(strNick)
  
  If InStr(Me.NickMode(strByNick), "o") = 0 Then
    ' ERR_CHANOPRIVSNEEDED
    conNick.SendData ":" & gstrServerName & " 482 " & strName & " :You're not channel operator"
    
    Exit Function
  End If
  
  conNick.SendData ":" & strByNick & " INVITE " & strNick & " " & strName
  
  Dim xmlInvite As clsSimpleXML
  Set xmlInvite = New clsSimpleXML
  
  xmlInvite.Tag = strNick
  
  colInvited.Add xmlInvite, xmlInvite.Tag
End Function

Public Function JoinNick(strNick As String)
  Dim conNick As clsConnect
  Set conNick = GetConnect(strNick)
  
  If (InStr(strModes, "i") > 0) And Not InvitedNick(strNick) Then
    ' ERR_INVITEONLYCHAN
    conNick.SendData ":" & gstrServerName & " 473 " & strName & " :Cannot join channel (+i)"
    
    Exit Function
  End If
  
  If BannedNick(strNick) Then
    ' ERR_BANNEDFROMCHAN
    conNick.SendData ":" & gstrServerName & " 474 " & strName & " :Cannot join channel (+b)"
    
    Exit Function
  End If
  
  
  conNick.SendData ":" & conNick.FullIdent & " JOIN " & strName
  
  Dim xmlNick As clsSimpleXML
  Set xmlNick = New clsSimpleXML
  
  xmlNick.Tag = strNick
  xmlNick.Value("FullIdent") = conNick.FullIdent
  
  If colMembers.Count = 0 Then
    xmlNick.Value("Modes") = "o"
  End If
  
  ' ??
  SendData ":" & xmlNick.Value("FullIdent") & " JOIN " & strName
  
  colMembers.Add xmlNick, xmlNick.Tag
  
  SetTopic xmlNick.Tag, ""
  
  Dim xmlSend As clsSimpleXML
  
  Dim lngIndex As Long
  lngIndex = 1
  
  Dim strNicks As String
  
  Do Until lngIndex > colMembers.Count
    Set xmlSend = colMembers(lngIndex)
    
    Select Case True
      Case (InStr(xmlSend.Value("Modes"), "o") > 0):
        strNicks = strNicks & "@" & xmlSend.Tag
      Case (InStr(xmlSend.Value("Modes"), "V") > 0):
        strNicks = strNicks & "+" & xmlSend.Tag
      Case Else:
        strNicks = strNicks & xmlSend.Tag
    End Select
    
    strNicks = strNicks & " "
  
    lngIndex = lngIndex + 1
  Loop
  
  strNicks = Trim(strNicks)
  
  ' RPL_NAMREPLY
  conNick.SendData ":" & gstrServerName & " 353 " & strNick & " " & strName & " :" & strNicks
  ' RPL_ENDOFNAMES
  conNick.SendData ":" & gstrServerName & " 366 " & strNick & " " & strName & " :End of /NAMES list"
  
  DebugPrint "[" & strDebugGUID & " |JOIN  ] " & strNick
End Function

Public Function PartNick(strNick As String, strReason As String)
  Dim xmlNick As clsSimpleXML

  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colMembers.Count
    Set xmlNick = colMembers(lngIndex)
    
    If LCase(xmlNick.Tag) = LCase(strNick) Then
      ' ??
      SendData ":" & xmlNick.Value("FullIdent") & " PART " & strName & " :" & strReason, True
    
      colMembers.Remove lngIndex
      
      DebugPrint "[" & strDebugGUID & " |PART  ] " & strNick
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
  
  If colMembers.Count = 0 Then
    
  End If
End Function

Public Function HaveNick(strNick As String) As Boolean
  Dim xmlNick As clsSimpleXML

  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colMembers.Count
    Set xmlNick = colMembers(lngIndex)
    
    If (LCase(xmlNick.Tag) = LCase(strNick)) Or (LCase(xmlNick.Value("FullIdent")) = LCase(strNick)) Then
      HaveNick = True
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Private Function InvitedNick(strNick As String) As Boolean
  Dim xmlNick As clsSimpleXML

  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colInvited.Count
    Set xmlNick = colInvited(lngIndex)
    
    If LCase(xmlNick.Tag) = LCase(strNick) Then
      InvitedNick = True
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Private Function BannedNick(strNick As String) As Boolean
  Dim xmlNick As clsSimpleXML

  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colBanned.Count
    Set xmlNick = colBanned(lngIndex)
    
    If (xmlNick.Tag Like strNick) And (Len(xmlNick.Tag) > 0) Then
      BannedNick = True
      
      Exit Do
    End If
    
    lngIndex = lngIndex + 1
  Loop
End Function

Public Function SendData(strData As String, Optional bolParting As Boolean = False)
  Dim xmlNick As clsSimpleXML
  Dim conNick As clsConnect

  Dim lngIndex As Long
  lngIndex = 1
  
  Dim strTempNick As String
  Dim bolSend As Boolean
  
  Do Until lngIndex > colMembers.Count
    Set xmlNick = colMembers(lngIndex)
    Set conNick = GetConnect(xmlNick.Tag)
    
    If (conNick Is Nothing) And Not bolParting Then
      PartNick xmlNick.Tag, "Client disconnect"
    Else
      bolSend = True
    
      If (Left(strData, 1) = ":") And (InStr(strData, "PRIVMSG") > 0) Then
        strTempNick = Right(strData, Len(strData) - 1)
        strTempNick = Left(strTempNick, InStr(strTempNick, " ") - 1)
        
        If (LCase(strTempNick) = LCase(xmlNick.Tag)) Or (LCase(strTempNick) = LCase(xmlNick.Value("FullIdent"))) Then
          bolSend = False
        End If
      End If
    
      If bolSend Then
        conNick.SendData strData
      End If
      
      lngIndex = lngIndex + 1
    End If
  Loop
End Function

Public Function SetTopic(strNick As String, strNewTopic As String)
  Dim conNick As clsConnect
  Set conNick = GetConnect(strNick)
  
  If conNick Is Nothing Then
    Exit Function
  End If
  
  If Len(strNewTopic) = 0 Then
    If Len(strTopic) = 0 Then
      ' RPL_NOTOPIC
      conNick.SendData ":" & gstrServerName & " 331 " & conNick.Nick & " " & strName & " :No topic is set"
    Else
      ' RPL_TOPIC
      conNick.SendData ":" & gstrServerName & " 332 " & conNick.Nick & " " & strName & " :" & strTopic
      ' ??
      conNick.SendData ":" & gstrServerName & " 333 " & conNick.Nick & " " & strName & " " & strTopicSetBy & " " & lngTopicSetTime
    End If
    
    Exit Function
  End If
  
  If Not HaveNick(strNick) Then
    ' ERR_NOTONCHANNEL
    conNick.SendData ":" & gstrServerName & " 442 " & strName & " :You're not on that channel"
    
    Exit Function
  End If
  
  Dim xmlNick As clsSimpleXML
  Set xmlNick = colMembers(strNick)
  
  Select Case True
    Case (InStr(xmlNick.Value("Modes"), "o") = 0):
      ' ERR_CHANOPRIVSNEEDED
      conNick.SendData ":" & gstrServerName & " 482 " & strName & " :You're not channel operator"
    Case Else:
      strTopic = strNewTopic
      strTopicSetBy = conNick.FullIdent
      lngTopicSetTime = UnixTime
      
      ' RPL_TOPIC
      SendData ":" & gstrServerName & " 332 " & conNick.Nick & " " & strName & " :" & strTopic
      ' ??
      SendData ":" & gstrServerName & " 333 " & conNick.Nick & " " & strName & " " & strTopicSetBy & " " & lngTopicSetTime
      
      DebugPrint "[" & strDebugGUID & " |TOPIC ] " & strTopic
  End Select
End Function

Public Function SendBanList(strNick As String)
  Dim conSend As clsConnect
  Set conSend = GetConnect(strNick)
  
  Dim xmlBanned As clsSimpleXML
  Dim strBanned As String
  
  Dim lngIndex As Long
  lngIndex = 1
  
  Do Until lngIndex > colBanned.Count
    Set xmlBanned = colBanned(lngIndex)
    strBanned = xmlBanned.Tag
  
    If Len(strBanned) > 0 Then
      ' RPL_BANLIST
      conSend.SendData ":" & gstrServerName & " 367 " & strNick & " " & strName & " " & strBanned
    End If
    
    lngIndex = lngIndex + 1
  Loop
  
  ' RPL_ENDOFBANLIST
  conSend.SendData ":" & gstrServerName & " 368 " & strNick & " " & strName & " :End of Channel Ban List"
End Function
