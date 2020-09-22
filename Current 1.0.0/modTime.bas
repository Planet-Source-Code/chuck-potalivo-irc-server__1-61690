Attribute VB_Name = "modTime"
Option Explicit

Public Function UnixTime() As Long
  UnixTime = DateDiff("s", DateValue("1/1/1970"), Now)
End Function

