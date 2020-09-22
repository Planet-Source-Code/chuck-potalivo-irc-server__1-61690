Attribute VB_Name = "modServer"
Option Explicit

Private Const cstrMOTDFilename    As String = "Config\MOTD.txt"

Public Function SendWelcome(strNick As String)
  Dim conNick As clsConnect
  Set conNick = GetConnect(strNick)

  ' RPL_WELCOME
  conNick.SendData ":" & gstrServerName & " 001 " & conNick.Nick & " :Welcome to the " & gstrNetworkName & ", " & conNick.FullIdent
  ' RPL_YOURHOST
  conNick.SendData ":" & gstrServerName & " 002 " & conNick.Nick & " :Your host is " & gstrServerName & ", running " & App.Title & " version " & gstrServerVersion
  ' RPL_CREATED
  conNick.SendData ":" & gstrServerName & " 003 " & conNick.Nick & " :This server was created " & gstrServerStartTime
  ' RPL_MYINFO
'  conNick.SendData ":" & gstrServerName & " 004 " & conNick.Nick & " " & gstrServerName & " " & gstrServerVersion & " iswo psitmn"
  ' RPL_VERSION
  conNick.SendData ":" & gstrServerName & " 351 " & conNick.Nick & " :" & gstrServerVersion & " " & gstrServerName & " :" & gstrServerComments
'  conNick.SendData ":" & gstrServerName & " 351 " & App.Title & "|" & gstrServerVersion & " " & gstrServerName & " :" & gstrServerComments

  DebugPrint "[" & conNick.DebugGUID & " |AUTH  ] Server Welcome Message sent"
End Function

Public Function SendMOTD(strNick As String)
  Dim strFilename As String
  strFilename = App.Path & "\" & cstrMOTDFilename
  
  Dim conNick As clsConnect
  Set conNick = GetConnect(strNick)
  
  ' RPL_MOTDSTART
  conNick.SendData ":" & gstrServerName & " 375 :Message of the day"
  
  ' RPL_MOTD
  conNick.SendData ":" & gstrServerName & " 372 :Begining of MOTD command"
  
  If Len(Dir(strFilename, vbNormal)) > 0 Then
    Dim intFNum As String
    intFNum = FreeFile
    
    Open strFilename For Input As intFNum
    
    Dim strLine As String
    
    Do Until EOF(intFNum)
      Line Input #intFNum, strLine
      
      ' RPL_MOTD
      conNick.SendData ":" & gstrServerName & " 372 :  " & strLine
    Loop
    
    Close #intFNum
  Else
    ' ERR_NOMOTD
    conNick.SendData ":" & gstrServerName & " 422 :MOTD File is missing"
  End If

  ' RPL_ENDOFMOTD
  conNick.SendData ":" & gstrServerName & " 376 :End of MOTD command"

  DebugPrint "[" & conNick.DebugGUID & " |AUTH  ] MOTD sent"
End Function
