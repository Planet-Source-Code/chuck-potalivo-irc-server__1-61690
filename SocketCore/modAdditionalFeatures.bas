Attribute VB_Name = "modAdditionalFeatures"
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Function GetDefaultIP() As String
    '----------------------------------------------------
    'pointer to HOSTENT structure returned by
    'the gethostbyname function
    Dim lngPtrToHOSTENT As Long
    '
    'structure which stores all the host info
    Dim udtHostent      As HOSTENT
    '
    'pointer to the IP address' list
    Dim lngPtrToIP      As Long
    '
    'byte array that contains elemets of an IP address
    Dim arrIpAddress()  As Byte
    '
    'result IP address string
    Dim strIpAddress    As String
    '
    'buffer string to receive the local system host name
    Dim strHostName As String * 256
    '
    Dim strLocalHost As String
    '
    'value returned by the gethostname function
    Dim lngRetVal As Long
    '----------------------------------------------------
    '
    'Get the local host name
    lngRetVal = gethostname(strHostName, 256)
    '
    If lngRetVal = SOCKET_ERROR Then
        Err.Raise Err.LastDllError, "modAdditionalFeatures.GetDefaultIP", GetErrorDescription(Err.LastDllError)
        Exit Function
    End If
    '
    strLocalHost = Left(strHostName, InStr(1, strHostName, Chr(0)) - 1)
    '
    'Call the gethostbyname Winsock API function
    'to get pointer to the HOSTENT structure
    lngPtrToHOSTENT = gethostbyname(Trim$(strLocalHost))
    '
    'Check the lngPtrToHOSTENT value
    If lngPtrToHOSTENT = 0 Then
        '
        'If the gethostbyname function has returned 0
        'the function execution is failed. To get
        'error description call the ShowErrorMsg
        'subroutine
        '
        Err.Raise Err.LastDllError, "modAdditionalFeatures.GetDefaultIP", GetErrorDescription(Err.LastDllError)
        '
    Else
        '
        'The gethostbyname function has found the address
        '
        'Copy retrieved data to udtHostent structure
        RtlMoveMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)
        '
        'Now udtHostent.hAddrList member contains
        'an array of IP addresses
        '
        'Get a pointer to the first address
        RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
        '
        Do Until lngPtrToIP = 0
            '
            'Prepare the array to receive IP address values
            ReDim arrIpAddress(1 To udtHostent.hLength)
            '
            'move IP address values to the array
            RtlMoveMemory arrIpAddress(1), lngPtrToIP, udtHostent.hLength
            '
            Dim I As Long
            '
            'build string with IP address
            For I = 1 To udtHostent.hLength
                strIpAddress = strIpAddress & arrIpAddress(I) & "."
            Next
            '
            'remove the last dot symbol
            strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
            '
            'Check to see if this is the default IP
            If strIpAddress <> "0.0.0.0" Then
              GetDefaultIP = strIpAddress
              
              Exit Do
            End If
            '
            'Clear the buffer
            strIpAddress = ""
            '
            'Get pointer to the next address
            udtHostent.hAddrList = udtHostent.hAddrList + LenB(udtHostent.hAddrList)
            RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
            '
         Loop
        '
    End If
    '
End Function

Public Function GetDefaultHostName() As String
    '----------------------------------------------------
    'pointer to HOSTENT structure returned by
    'the gethostbyname function
    Dim lngPtrToHOSTENT As Long
    '
    'structure which stores all the host info
    Dim udtHostent      As HOSTENT
    '
    'pointer to the IP address' list
    Dim lngPtrToIP      As Long
    '
    'byte array that contains elemets of an IP address
    Dim arrIpAddress()  As Byte
    '
    'result IP address string
    Dim strIpAddress    As String
    '
    'buffer string to receive the local system host name
    Dim strHostName As String * 256
    '
    Dim strLocalHost As String
    '
    'value returned by the gethostname function
    Dim lngRetVal As Long
    '----------------------------------------------------
    '
    'Get the local host name
    lngRetVal = gethostname(strHostName, 256)
    '
    If lngRetVal = SOCKET_ERROR Then
        Err.Raise Err.LastDllError, "modAdditionalFeatures.GetDefaultIP", GetErrorDescription(Err.LastDllError)
        Exit Function
    End If
    '
    strLocalHost = Left(strHostName, InStr(1, strHostName, Chr(0)) - 1)
    '
    GetDefaultHostName = strLocalHost
End Function

