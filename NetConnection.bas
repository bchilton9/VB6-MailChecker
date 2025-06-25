Attribute VB_Name = "NetConnection"
Option Explicit

Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412
Public Const ERROR_SUCCESS = 0&

Public Type RasEntryName
    dwSize As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
End Type

Public Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Public Declare Function RasEnumConnections Lib _
"rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As _
Any, lpcb As Long, lpcConnections As Long) As Long

Public Declare Function RasHangUp Lib "rasapi32.dll" Alias _
"RasHangUpA" (ByVal hRasConn As Long) As Long
Public gstrISPName As String

Public ReturnCode As Long
Public Declare Function InternetGetConnectedState _
    Lib "wininet.dll" (ByRef lpdwFlags As Long, _
    ByVal dwReserved As Long) As Long
    'Local system uses a modem to connect to
    '     the Internet.
    Public Const INTERNET_CONNECTION_MODEM As Long = &H1
    'Local system uses a LAN to connect to t
    '     he Internet.
    Public Const INTERNET_CONNECTION_LAN As Long = &H2
    'Local system uses a proxy server to con
    '     nect to the Internet.
    Public Const INTERNET_CONNECTION_PROXY As Long = &H4
    'No longer used.
    Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
    Public Const INTERNET_RAS_INSTALLED As Long = &H10
    Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
    Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
    'InternetGetConnectedState wrapper funct
    '     ions


Public Function IsNetConnectViaLAN() As Boolean
    Dim dwflags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwflags, 0&)
    'return True if the flags indicate a LAN
    '     connection
    IsNetConnectViaLAN = dwflags And INTERNET_CONNECTION_LAN
End Function


Public Function IsNetConnectViaModem() As Boolean
    Dim dwflags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwflags, 0&)
    'return True if the flags indicate a mod
    '     em connection
    IsNetConnectViaModem = dwflags And INTERNET_CONNECTION_MODEM
End Function


Public Function IsNetConnectViaProxy() As Boolean
    Dim dwflags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwflags, 0&)
    'return True if the flags indicate a pro
    '     xy connection
    IsNetConnectViaProxy = dwflags And INTERNET_CONNECTION_PROXY
End Function


Public Function IsNetConnectOnline() As Boolean
    'no flags needed here - the API returns
    '     True
    'if there is a connection of any type
    IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
End Function


Public Function IsNetRASInstalled() As Boolean
    Dim dwflags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwflags, 0&)
    'return True if the falgs include RAS in
    '     stalled
    IsNetRASInstalled = dwflags And INTERNET_RAS_INSTALLED
End Function


Public Function GetNetConnectString() As String
    Dim dwflags As Long
    Dim msg As String
    'build a string for display


    If InternetGetConnectedState(dwflags, 0&) Then


        If dwflags And INTERNET_CONNECTION_CONFIGURED Then
            msg = msg & "You have a network connection configured." & vbCrLf
        End If


        If dwflags And INTERNET_CONNECTION_LAN Then
            msg = msg & "The local system connects To the Internet via a LAN"
        End If


        If dwflags And INTERNET_CONNECTION_PROXY Then
            msg = msg & ", and uses a proxy server. "
        Else: msg = msg & "."
        End If


        If dwflags And INTERNET_CONNECTION_MODEM Then
            msg = msg & "The local system uses a modem To connect to the Internet. "
        End If


        If dwflags And INTERNET_CONNECTION_OFFLINE Then
            msg = msg & "The connection is currently offline. "
        End If


        If dwflags And INTERNET_CONNECTION_MODEM_BUSY Then
            msg = msg & "The local system's modem is busy With a non-Internet connection. "
        End If


        If dwflags And INTERNET_RAS_INSTALLED Then
            msg = msg & "Remote Access Services are installed On this system."
        End If
    Else
        msg = "Not connected To the internet now."
    End If
    GetNetConnectString = msg
End Function



Public Sub HangUp()
Dim i As Long
Dim lpRasConn(255) As RasConn
Dim lpcb As Long
Dim lpcConnections As Long
Dim hRasConn As Long
lpRasConn(0).dwSize = RAS_RASCONNSIZE
lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
lpcConnections = 0
ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, _
lpcConnections)

If ReturnCode = ERROR_SUCCESS Then
    For i = 0 To lpcConnections - 1
        If Trim(ByteToString(lpRasConn(i).szEntryName)) _
            = Trim(gstrISPName) Then
            hRasConn = lpRasConn(i).hRasConn
            ReturnCode = RasHangUp(ByVal hRasConn)
        End If
    Next i
End If

End Sub

Public Function ByteToString(bytString() As Byte) As String
Dim i As Integer
ByteToString = ""
i = 0
While bytString(i) = 0&
ByteToString = ByteToString & Chr(bytString(i))
i = i + 1
Wend
End Function
