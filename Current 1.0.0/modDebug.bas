Attribute VB_Name = "modDebug"
Option Explicit

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Sub DebugPrint(strText As String)
  On Error Resume Next
  
  Dim strTimeStamp As String
  strTimeStamp = Format(Date & " " & Time, "MM/DD/YYYY HH:MM:SS")
 
  Dim itmConsole As ListItem
  Set itmConsole = frmMain.lstConsole.ListItems.Add(, , strTimeStamp)
  
  itmConsole.SubItems(1) = strText
  
  itmConsole.EnsureVisible
  itmConsole.Selected = True
  
  frmMain.lstConsole.Refresh

  Dim intFNum As Integer
  intFNum = FreeFile
  
  Dim strFilename As String
  strFilename = Replace(gstrServerStartTime, "/", "-") & ".txt"
  strFilename = Replace(strFilename, ":", ".")

  Open App.Path & "\Logs\" & strFilename For Append As intFNum

  Print #intFNum, strTimeStamp & "  " & strText

  Close #intFNum
  
  Dim telSend     As clsTelnet
  
  For Each telSend In colTelnet
    If telSend.ConsoleMode Then
      telSend.SendData strText & vbCrLf
    End If
  Next
End Sub


