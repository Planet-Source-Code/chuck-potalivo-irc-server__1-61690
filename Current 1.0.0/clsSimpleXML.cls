VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSimpleXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim strXML As String

Public Property Let XMLData(strValue As String)
  strXML = strValue
End Property

Public Property Get XMLData() As String
  XMLData = strXML
End Property

Public Function IsXML() As Boolean
  strXML = Trim(strXML)
  
  IsXML = (strXML Like "<*>")
End Function

Public Function ValueExist(strValue As String) As Boolean
  Dim strSearchValue As String

  strSearchValue = " " & Trim(strValue) & "="
  
  If InStr(LCase(strXML), LCase(strSearchValue)) = 0 Then
    ValueExist = False
  Else
    ValueExist = True
  End If
End Function

Public Property Get Value(strValue As String) As Variant
  Dim strSearchValue As String

  strSearchValue = " " & Trim(strValue) & "="
  
  If InStr(LCase(strXML), LCase(strSearchValue)) = 0 Then
    Value = ""
    
    Exit Property
  End If
  
  Dim intValueStart As Integer
  
  intValueStart = InStr(LCase(strXML), LCase(strSearchValue)) + Len(strSearchValue)
  
  Dim strTemp As String
  
  strTemp = Right(strXML, Len(strXML) - intValueStart)
  
  Dim intValueStop As Integer
  
  intValueStop = InStr(strTemp, """ ")
  
  If intValueStop = 0 Then
    intValueStop = InStr(strTemp, """>")
  End If
  
  intValueStop = intValueStop - 1
  
  Value = Left(strTemp, intValueStop)
End Property

Public Property Let Value(strValue As String, varValue As Variant)
  Dim strSearchValue As String

  strSearchValue = " " & Trim(strValue) & "="
  
  If InStr(LCase(strXML), LCase(strSearchValue)) = 0 Then
    strXML = Replace(strXML, ">", " " & strValue & "=""" & varValue & """>")
    strXML = Replace(strXML, """  ", """ ")
    
    Exit Property
  End If
  
  Dim intValueStart As Integer
  
  intValueStart = InStr(LCase(strXML), LCase(strSearchValue)) + Len(strSearchValue)
  
  Dim strTemp As String
  
  strTemp = Right(strXML, Len(strXML) - intValueStart)
  
  Dim intValueStop As Integer
  
  intValueStop = InStr(strTemp, """ ")
  
  If intValueStop = 0 Then
    intValueStop = InStr(strTemp, """>")
  End If
  
  intValueStop = intValueStop - 1
  
  Dim varCurrentValue As Variant
  
  varCurrentValue = Left(strTemp, intValueStop)
  
  Dim strSearchFor As String
  
  strSearchFor = strValue & "=""" & varCurrentValue & """"
  
  Dim strReplaceWith As String
  
  strReplaceWith = strValue & "=""" & varValue & """"
  
  strXML = Replace(strXML, strSearchFor, strReplaceWith)
End Property

Public Property Get Tag() As String
  Dim intTagStart As Integer
  
  intTagStart = InStr(strXML, "<")

  Dim intTagStop As Integer
  
  intTagStop = InStr(strXML, " ") - 1

  If intTagStop < 0 Then
    intTagStop = InStr(strXML, ">") - 1
  End If

  Dim strTemp As String
  
  strTemp = Left(strXML, intTagStop)
  strTemp = Right(strTemp, Len(strTemp) - intTagStart)
  
  Tag = strTemp
End Property

Public Property Let Tag(strValue As String)
  Dim intTagStart As Integer
  
  intTagStart = InStr(strXML, "<")

  Dim intTagStop As Integer
  
  intTagStop = InStr(strXML, " ") - 1
  
  If intTagStop < 0 Then
    intTagStop = InStr(strXML, ">") - 1
  End If

  Dim strTemp As String
  
  strTemp = Left(strXML, intTagStop)
  strTemp = Right(strTemp, Len(strTemp) - intTagStart)
  
  strXML = Replace(strXML, "<" & strTemp, "<" & strValue)
End Property

Private Sub Class_Initialize()
  strXML = "<DEFAULT>"
End Sub

