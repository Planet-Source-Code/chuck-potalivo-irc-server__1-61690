Attribute VB_Name = "modMain"
Option Explicit

Public colServers             As Collection
Public colChannels            As Collection
Public colConnects            As Collection

Public colTelnet              As Collection

Global gstrServerName         As String
Global gstrServerVersion      As String
Global gstrServerStartTime    As String
Global gstrServerComments     As String

Global gstrNetworkName        As String

Global glngPingInterval       As Long
Global glngSendInterval       As Long

Global glngMaxConnections     As Long
Global glngMaxJoins           As Long

Global glngTelnetPort         As Long
Global glngIRCPort            As Long

Global gbolShowSentData       As Boolean
Global gbolShowRcvdData       As Boolean

Private Const cstrConfigFile  As String = "Config\Config.xml"

Private Const cstrDefaultConfig As String = "<Config>" & vbCrLf & "</Config>"

Sub Main()
  On Error Resume Next

  gstrServerStartTime = Format(Date & " " & Time, "MM/DD/YYYY HH:MM:SS")
  gstrServerVersion = App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "000")
  
  Load frmMain
  
  frmMain.Show
  
  LoadSettings
  
  Dim strMain As String
  strMain = Space(20)
  
  LSet strMain = "MAIN " & gstrServerName
  
  Dim strDebugGUID As String
  strDebugGUID = GenerateGUID

  strDebugGUID = Replace(strDebugGUID, "-", "")
  strDebugGUID = Replace(strDebugGUID, "}", "")
  strDebugGUID = Replace(strDebugGUID, "{", "")
  strDebugGUID = Left(strDebugGUID, (Len(strDebugGUID) / 2) - 6)
  strDebugGUID = strMain & "|" & strDebugGUID
  
  DebugPrint "[" & strDebugGUID & " |INIT  ] Initializing IRC Server"
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Version:                  " & gstrServerVersion
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Server Name:              " & gstrServerName
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Network Name:             " & gstrNetworkName
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Max Connections:          " & Format(glngMaxConnections, "#,##0")
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Max Joins per Connection: " & Format(glngMaxJoins, "#,##0")
  DebugPrint "[" & strDebugGUID & " |INIT  ]   Thread ID:                " & GetCurrentThreadId
  
  
  Set colTelnet = New Collection
  
  Dim telNew As clsTelnet
  
  Set telNew = New clsTelnet
  telNew.Initialize glngTelnetPort, "127.0.0.1"
  colTelnet.Add telNew, telNew.GUID

  Set telNew = New clsTelnet
  telNew.Initialize glngTelnetPort
  colTelnet.Add telNew, telNew.GUID


  Set colServers = New Collection
  
  Dim lsnNew As clsListen
  
  Set lsnNew = New clsListen
  lsnNew.Initialize glngIRCPort, "127.0.0.1"
  colServers.Add lsnNew, lsnNew.GUID
  
  Set lsnNew = New clsListen
  lsnNew.Initialize glngIRCPort
  colServers.Add lsnNew, lsnNew.GUID
  
  
  Set colChannels = New Collection
  
  Dim chnNew As clsChannel
  
  Set chnNew = New clsChannel
  chnNew.Initialize "#help"
  colChannels.Add chnNew, chnNew.GUID
  
  Set chnNew = New clsChannel
  chnNew.Initialize "#test"
  colChannels.Add chnNew, chnNew.GUID
  
  
  Set colConnects = New Collection
End Sub

Public Function LoadSettings()
  Dim xmlLoad     As clsGoXML
  Set xmlLoad = New clsGoXML
  
  xmlLoad.Initialize pavAuto
  
  Dim strFilename As String
  strFilename = App.Path & "/" & cstrConfigFile
  
  If Dir(strFilename, vbNormal) = "" Then
    xmlLoad.OpenFromString cstrDefaultConfig, False
  
    xmlLoad.InsertNode "//Config", "Server"
    xmlLoad.InsertNode "//Config", "Console"
    xmlLoad.InsertNode "//Config", "Network"
    xmlLoad.InsertNode "//Config", "Client"
    xmlLoad.InsertNode "//Config", "Telnet"
  Else
    xmlLoad.OpenFromFile strFilename, False
  End If
  
  gstrServerName = xmlLoad.ReadAttribute("//Config/Server", "Name", "localhost")
  gstrServerComments = xmlLoad.ReadAttribute("//Config/Server", "Comments", "")

  gbolShowSentData = xmlLoad.ReadAttribute("//Config/Console", "ShowSentData", 1)
  gbolShowRcvdData = xmlLoad.ReadAttribute("//Config/Console", "ShowRcvdData", 1)

  gstrNetworkName = xmlLoad.ReadAttribute("Config/Network", "Name", "IRCServerNet")

  glngPingInterval = xmlLoad.ReadAttribute("//Config/Client", "PingInterval", 60000)
  glngSendInterval = xmlLoad.ReadAttribute("//Config/Client", "SendInterval", 100)

  glngMaxConnections = xmlLoad.ReadAttribute("//Config/Server", "MaxConnects", 10000)
  glngMaxJoins = xmlLoad.ReadAttribute("//Config/Client", "MaxJoins", 1000)

  glngTelnetPort = xmlLoad.ReadAttribute("//Config/Telnet", "Port", 23)
  glngIRCPort = xmlLoad.ReadAttribute("//Config/Server", "Port", 6667)
  
  xmlLoad.Save strFilename
  
  Set xmlLoad = Nothing
End Function

Public Function SaveSettings()
  Dim xmlSave     As clsGoXML
  Set xmlSave = New clsGoXML
  
  xmlSave.Initialize pavAuto
  
  Dim strFilename As String
  strFilename = App.Path & "/" & cstrConfigFile
  
  xmlSave.OpenFromFile strFilename, False
  
  xmlSave.WriteAttribute "//Config/Server", "Name", gstrServerName
  xmlSave.WriteAttribute "//Config/Server", "Comments", gstrServerComments
  
  xmlSave.WriteAttribute "//Config/Console", "ShowSentData", gbolShowSentData
  xmlSave.WriteAttribute "//Config/Console", "ShowRcvdData", gbolShowRcvdData
  
  xmlSave.WriteAttribute "//Config/Network", "Name", gstrNetworkName
  
  xmlSave.WriteAttribute "//Config/Client", "PingInterval", glngPingInterval
  xmlSave.WriteAttribute "//Config/Client", "SendInterval", glngSendInterval
  
  xmlSave.WriteAttribute "//Config/Server", "MaxConnets", glngMaxConnections
  xmlSave.WriteAttribute "//Config/Client", "MaxJoins", glngMaxJoins
  
  xmlSave.WriteAttribute "//Config/Telnet", "Port", glngTelnetPort
  xmlSave.WriteAttribute "//Config/Server", "Port", glngIRCPort
  
  xmlSave.Save strFilename
  
  Set xmlSave = Nothing
End Function
