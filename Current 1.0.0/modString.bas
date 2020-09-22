Attribute VB_Name = "modString"
Option Explicit

Public Function SmartSplit(ByVal strExpression As String, Optional strDelimeter As String = " ") As Variant
  Dim arySplit() As String
  Dim lngSplit As Long
  
  Dim strSplit As String
  
  Do Until Len(strExpression) = 0
    If (Left(strExpression, 1) = ":") And (lngSplit > 0) Then
      strSplit = Right(strExpression, Len(strExpression) - 1)
    
      ReDim Preserve arySplit(0 To lngSplit) As String
      arySplit(lngSplit) = strSplit
      
      strExpression = ""
    ElseIf (InStr(strExpression, " ") > InStr(strExpression, """")) And (InStr(strExpression, """") > 0) Then
      strSplit = Right(strExpression, Len(strExpression) - InStr(strExpression, """"))
      strSplit = Left(strSplit, InStr(strSplit, """") - 1)
      
      ReDim Preserve arySplit(0 To lngSplit) As String
      arySplit(lngSplit) = strSplit
      
      strExpression = Right(strExpression, Len(strExpression) - InStr(strExpression, """"))
      strExpression = Right(strExpression, Len(strExpression) - InStr(strExpression, """"))
      strExpression = LTrim(strExpression)
    ElseIf (InStr(strExpression, " ") > 0) Then
      strSplit = Left(strExpression, InStr(strExpression, " ") - 1)
      
      ReDim Preserve arySplit(0 To lngSplit) As String
      arySplit(lngSplit) = strSplit
      
      strExpression = Right(strExpression, Len(strExpression) - InStr(strExpression, " "))
    Else
      strSplit = strExpression
    
      ReDim Preserve arySplit(0 To lngSplit) As String
      arySplit(lngSplit) = strSplit
      
      strExpression = ""
    End If
    
    lngSplit = lngSplit + 1
    
'    DebugPrint "[SPLIT]  '" & strSplit & "'"
  Loop
  
  SmartSplit = arySplit
End Function
