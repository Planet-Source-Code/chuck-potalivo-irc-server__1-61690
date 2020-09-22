VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long

Private Const LOGON32_LOGON_INTERACTIVE   As Long = 2
Private Const LOGON32_LOGON_NETWORK       As Long = 3
Private Const LOGON32_PROVIDER_DEFAULT    As Long = 0
Private Const LOGON32_PROVIDER_WINNT50    As Long = 3
Private Const LOGON32_PROVIDER_WINNT40    As Long = 2
Private Const LOGON32_PROVIDER_WINNT35    As Long = 1

Public Function VerifyLogin(sUser As String, sDomain As String, sPassword As String) As Boolean
  Dim token As Long
  
  VerifyLogin = LogonUser(sUser, sDomain, sPassword, LOGON32_LOGON_NETWORK, LOGON32_PROVIDER_DEFAULT, token)
End Function

Public Function GetCurrentUser() As String
  Dim strName As String
  strName = String$(255, 0)
  
  Dim x As Long
  x = GetUserName(strName, Len(strName))
    
  GetCurrentUser = strName
End Function
