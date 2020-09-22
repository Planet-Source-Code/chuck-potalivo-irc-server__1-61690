Attribute VB_Name = "modValidation"
Option Explicit

Public Function IsValidNick(strNick As String) As Boolean
  Select Case True
    Case (InStr(strNick, "*") > 0):
    Case (InStr(strNick, " ") > 0):
    Case (InStr(strNick, "!") > 0):
    Case (InStr(strNick, "@") > 0):
    Case (InStr(strNick, "#") > 0):
    Case (Len(strNick) > clngNickMaxLen):
    Case (strNick Like "NickServ"):
    Case (strNick Like "OperServ"):
    Case Else:
      IsValidNick = True
  End Select
End Function

Public Function IsValidChan(strChan As String) As Boolean
  Select Case True
    Case (Left(strChan, 1) <> "#"):
    Case Else:
      IsValidChan = True
  End Select
End Function
