Attribute VB_Name = "Winsock2"
Option Explicit

'
' ---------------------------------------------------------------------------------
' File...........: mWinsock2.bas
' Author.........: Will Barden
' Created........: 02/05/03
' Modified.......: 09/05/03
' Version........: 1.0
' Website........: http://www.WinsockVB.com
' Contact........: admin@winsockvb.com
'
' Port of necessary Winsock2 declares, consts, types etc.. Will handle straight
' blocking I/O, WSAAsyncSelect and WSAEventSelect under both TCP/IP and UDP/IP.
' Will also handle sending ICMP echos (a ping) to check if a host is alive. Has
' some helper functions at the bottom, prefixed with "vb".
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
' Constants.
' ------------------------------------------------------------------------------
'
' Winsock version constants.
Public Const WINSOCK_V1_1  As Long = &H101
Public Const WINSOCK_V2_2  As Long = &H202
'
' Length of fields within the WSADATA structure.
Public Const WSADESCRIPTION_LEN  As Long = 256
Public Const WSASYS_STATUS_LEN   As Long = 128
'
' For socket handle errors, and bas returns from APIs.
Public Const ERROR_SUCCESS    As Long = 0
Public Const SOCKET_ERROR     As Long = -1
Public Const INVALID_SOCKET   As Long = SOCKET_ERROR
'
' Internet addresses.
Public Const INADDR_ANY          As Long = &H0
Public Const INADDR_LOOPBACK     As Long = &H7F000001
Public Const INADDR_BROADCAST    As Long = &HFFFFFFFF
Public Const INADDR_NONE         As Long = &HFFFFFFFF
'
' Maximum backlog when calling listen().
Public Const SOMAXCONN  As Long = 5
'
' Messages send with WSAAsyncSelect().
Public Const FD_READ       As Long = &H1
Public Const FD_WRITE      As Long = &H2
Public Const FD_OOB        As Long = &H4
Public Const FD_ACCEPT     As Long = &H8
Public Const FD_CONNECT    As Long = &H10
Public Const FD_CLOSE      As Long = &H20
'
' Used with shutdown().
Public Const SD_RECEIVE    As Long = &H0
Public Const SD_SEND       As Long = &H1
Public Const SD_BOTH       As Long = &H2
'
' Winsock error constants.
Public Const WSABASEERR          As Long = 10000
Public Const WSAEINTR            As Long = WSABASEERR + 4
Public Const WSAEBADF            As Long = WSABASEERR + 9
Public Const WSAEACCES           As Long = WSABASEERR + 13
Public Const WSAEFAULT           As Long = WSABASEERR + 14
Public Const WSAEINVAL           As Long = WSABASEERR + 22
Public Const WSAEMFILE           As Long = WSABASEERR + 24
Public Const WSAEWOULDBLOCK      As Long = WSABASEERR + 35
Public Const WSAEINPROGRESS      As Long = WSABASEERR + 36
Public Const WSAEALREADY         As Long = WSABASEERR + 37
Public Const WSAENOTSOCK         As Long = WSABASEERR + 38
Public Const WSAEDESTADDRREQ     As Long = WSABASEERR + 39
Public Const WSAEMSGSIZE         As Long = WSABASEERR + 40
Public Const WSAEPROTOTYPE       As Long = WSABASEERR + 41
Public Const WSAENOPROTOOPT      As Long = WSABASEERR + 42
Public Const WSAEPROTONOSUPPORT  As Long = WSABASEERR + 43
Public Const WSAESOCKTNOSUPPORT  As Long = WSABASEERR + 44
Public Const WSAEOPNOTSUPP       As Long = WSABASEERR + 45
Public Const WSAEPFNOSUPPORT     As Long = WSABASEERR + 46
Public Const WSAEAFNOSUPPORT     As Long = WSABASEERR + 47
Public Const WSAEADDRINUSE       As Long = WSABASEERR + 48
Public Const WSAEADDRNOTAVAIL    As Long = WSABASEERR + 49
Public Const WSAENETDOWN         As Long = WSABASEERR + 50
Public Const WSAENETUNREACH      As Long = WSABASEERR + 51
Public Const WSAENETRESET        As Long = WSABASEERR + 52
Public Const WSAECONNABORTED     As Long = WSABASEERR + 53
Public Const WSAECONNRESET       As Long = WSABASEERR + 54
Public Const WSAENOBUFS          As Long = WSABASEERR + 55
Public Const WSAEISCONN          As Long = WSABASEERR + 56
Public Const WSAENOTCONN         As Long = WSABASEERR + 57
Public Const WSAESHUTDOWN        As Long = WSABASEERR + 58
Public Const WSAETOOMANYREFS     As Long = WSABASEERR + 59
Public Const WSAETIMEDOUT        As Long = WSABASEERR + 60
Public Const WSAECONNREFUSED     As Long = WSABASEERR + 61
Public Const WSAELOOP            As Long = WSABASEERR + 62
Public Const WSAENAMETOOLONG     As Long = WSABASEERR + 63
Public Const WSAEHOSTDOWN        As Long = WSABASEERR + 64
Public Const WSAEHOSTUNREACH     As Long = WSABASEERR + 65
Public Const WSAENOTEMPTY        As Long = WSABASEERR + 66
Public Const WSAEPROCLIM         As Long = WSABASEERR + 67
Public Const WSAEUSERS           As Long = WSABASEERR + 68
Public Const WSAEDQUOT           As Long = WSABASEERR + 69
Public Const WSAESTALE           As Long = WSABASEERR + 70
Public Const WSAEREMOTE          As Long = WSABASEERR + 71
Public Const WSASYSNOTREADY      As Long = WSABASEERR + 91
Public Const WSAVERNOTSUPPORTED  As Long = WSABASEERR + 92
Public Const WSANOTINITIALISED   As Long = WSABASEERR + 93
Public Const WSAHOST_NOT_FOUND   As Long = WSABASEERR + 1001
'
' Winsock 2 extensions.
Public Const WSA_IO_PENDING         As Long = 997
Public Const WSA_IO_INCOMPLETE      As Long = 996
Public Const WSA_INVALID_HANDLE     As Long = 6
Public Const WSA_INVALID_PARAMETER  As Long = 87
Public Const WSA_NOT_ENOUGH_MEMORY  As Long = 8
Public Const WSA_OPERATION_ABORTED  As Long = 995

Public Const WSA_WAIT_FAILED           As Long = -1
Public Const WSA_WAIT_EVENT_0          As Long = 0
Public Const WSA_WAIT_IO_COMPLETION    As Long = &HC0
Public Const WSA_WAIT_TIMEOUT          As Long = &H102
Public Const WSA_INFINITE              As Long = -1
'
' Max size of event handle array when calling WSAWaitForMultipleEvents().
Public Const WSA_MAXIMUM_WAIT_EVENTS   As Long = 64
'
' Size of WSANETWORKEVENTS.iErrorCode[] array.
Public Const FD_MAX_EVENTS    As Long = 10
'
' Used to refer to particular elements of the WSANETWORKEVENTS.iErrorCodes[].
Public Const FD_READ_BIT                     As Long = 0
Public Const FD_WRITE_BIT                    As Long = 1
Public Const FD_OOB_BIT                      As Long = 2
Public Const FD_ACCEPT_BIT                   As Long = 3
Public Const FD_CONNECT_BIT                  As Long = 4
Public Const FD_CLOSE_BIT                    As Long = 5
Public Const FD_QOS_BIT                      As Long = 6
Public Const FD_GROUP_QOS_BIT                As Long = 7
Public Const FD_ROUTING_INTERFACE_CHANGE_BIT As Long = 8
Public Const FD_ADDRESS_LIST_CHANGE_BIT      As Long = 9
'
' ------------------------------------------------------------------------------
' Enumerations.
' ------------------------------------------------------------------------------
'
' Used with socket().
Public Enum Protocols
   IPPROTO_IP = 0
   IPPROTO_ICMP = 1
   IPPROTO_GGP = 2
   IPPROTO_TCP = 6
   IPPROTO_PUP = 12
   IPPROTO_UDP = 17
   IPPROTO_IDP = 22
   IPPROTO_ND = 77
   IPPROTO_RAW = 255
   IPPROTO_MAX = 256
End Enum
'
' Used with socket().
Public Enum SocketTypes
   SOCK_STREAM = 1
   SOCK_DGRAM = 2
   SOCK_RAW = 3
   SOCK_RDM = 4
   SOCK_SEQPACKET = 5
End Enum
'
' Used with socket().
Public Enum AddressFamilies
   AF_UNSPEC = 0
   AF_UNIX = 1
   AF_INET = 2
   AF_IMPLINK = 3
   AF_PUP = 4
   AF_CHAOS = 5
   AF_NS = 6
   AF_IPX = 6
   AF_ISO = 7
   AF_OSI = 7
   AF_ECMA = 8
   AF_DATAKIT = 9
   AF_CCITT = 10
   AF_SNA = 11
   AF_DECNET = 12
   AF_DLI = 13
   AF_LAT = 14
   AF_HYLINK = 15
   AF_APPLETALK = 16
   AF_NETBIOS = 17
   AF_MAX = 18
End Enum
'
' ------------------------------------------------------------------------------
' Types.
' ------------------------------------------------------------------------------
'
' To initialize Winsock.
Public Type WSADATA
   wVersion                               As Integer
   wHighVersion                           As Integer
   szDescription(WSADESCRIPTION_LEN + 1)  As Byte
   szSystemstatus(WSASYS_STATUS_LEN + 1)  As Byte
   iMaxSockets                            As Integer
   iMaxUpdDg                              As Integer
   lpVendorInfo                           As Long
End Type
'
' Basic IPv4 addressing structures.
Public Type in_addr
   s_addr   As Long
End Type
'
Public Type sockaddr_in
   sin_family        As Integer
   sin_port          As Integer
   sin_addr          As in_addr
   sin_zero(0 To 7)  As Byte
End Type
'
' Used with name resolution functions.
Public Type hostent
   h_name         As Long
   h_aliases      As Long
   h_addrtype     As Integer
   h_length       As Integer
   h_addr_list    As Long
End Type
'
' Used with WSAEnumNetworkEvents().
Public Type WSANETWORKEVENTS
    lNetworkEvents               As Long
    iErrorCode(FD_MAX_EVENTS)    As Integer
End Type
'
' Used when sending ICMP echos (pings).
Public Type IP_OPTION_INFORMATION
    TTL           As Byte
    Tos           As Byte
    Flags         As Byte
    OptionsSize   As Long
    OptionsData   As String * 128
End Type
'
Public Type IP_ECHO_REPLY
    Address(0 To 3)  As Byte
    Status           As Long
    RoundTripTime    As Long
    DataSize         As Integer
    Reserved         As Integer
    data             As Long
    Options          As IP_OPTION_INFORMATION
End Type
'
' ------------------------------------------------------------------------------
' APIs.
' ------------------------------------------------------------------------------
'
' DLL handling functions.
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Public Declare Function WSASetLastError Lib "ws2_32.dll" (ByVal err As Long) As Long
'
' Resolution functions.
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
'
' Conversion functions.
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal laddr As Long) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer
'
' Socket functions.
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As AddressFamilies, ByVal stype As SocketTypes, ByVal protocol As Protocols) As Long
'
Public Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
Public Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As sockaddr_in, ByRef addrlen As Long) As Long
Public Declare Function connect Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
'
Public Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long) As Long
Public Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef toaddr As sockaddr_in, ByVal tolen As Long) As Long
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef fromaddr As sockaddr_in, ByRef fromlen As Long) As Long
'
Public Declare Function shutdown Lib "ws2_32.dll" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
'
' I/O model functions.
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Integer, ByVal lEvent As Long) As Long
'
Public Declare Function WSACreateEvent Lib "ws2_32.dll" () As Long
Public Declare Function WSAEventSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hEventObject As Long, ByVal lNetworkEvents As Long) As Long
Public Declare Function WSAResetEvent Lib "ws2_32.dll" (ByVal hEvent As Long) As Long
Public Declare Function WSASetEvent Lib "ws2_32.dll" (ByVal hEvent As Long) As Long
Public Declare Function WSACloseEvent Lib "ws2_32.dll" (ByVal hEvent As Long) As Long
Public Declare Function WSAWaitForMultipleEvents Lib "ws2_32.dll" (ByVal cEvents As Long, ByRef lphEvents As Long, ByVal fWaitAll As Boolean, ByVal dwTimeout As Long, ByVal fAlertable As Boolean) As Long
Public Declare Function WSAEnumNetworkEvents Lib "ws2_32.dll" (ByVal s As Long, ByVal hEvent As Long, ByRef lpNetworkEvents As WSANETWORKEVENTS) As Long
'
' ICMP functions.
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Public Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean
'
' Other general Win32 APIs.
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'
' ------------------------------------------------------------------------------
' Helper methods.
' ------------------------------------------------------------------------------
'
Public Function vbGetLastError(Optional lngErrorCode As Long = 0) As String
   '
Dim lngNum As Long
Dim strRet As String
   '
   ' Return a useful description of the last winsock error.
   If (lngErrorCode) Then
      lngNum = lngErrorCode
   Else
      lngNum = WSAGetLastError()
   End If
   '
   Select Case lngNum
      '
      ' Winsock errors.
      Case WSAEINTR
         strRet = "interrupted function call"
      Case WSAEACCES
         strRet = "permission denied"
      Case WSAEFAULT
         strRet = "invalid address"
      Case WSAEINVAL
         strRet = "invalid argument"
      Case WSAEMFILE
         strRet = "too many files open"
      Case WSAEWOULDBLOCK
         strRet = "function call would block"
      Case WSAEINPROGRESS
         strRet = "blocking call already in progress"
      Case WSAEALREADY
         strRet = "operation already in progress"
      Case WSAENOTSOCK
         strRet = "not a valid socket descriptor"
      Case WSAEDESTADDRREQ
         strRet = "destination address required"
      Case WSAEMSGSIZE
         strRet = "message is too long"
      Case WSAEPROTOTYPE
         strRet = "protocol wrong type for socket"
      Case WSAENOPROTOOPT
         strRet = "bad protocol option"
      Case WSAEPROTONOSUPPORT
         strRet = "protocol not supported"
      Case WSAESOCKTNOSUPPORT
         strRet = "socket type not supported"
      Case WSAEOPNOTSUPP
         strRet = "operation not supported"
      Case WSAEPFNOSUPPORT
         strRet = "protocol family not supported"
      Case WSAEAFNOSUPPORT
         strRet = "address family not supported by protocol"
      Case WSAEADDRINUSE
         strRet = "address in use"
      Case WSAEADDRNOTAVAIL
         strRet = "address is not available"
      Case WSAENETDOWN
         strRet = "network is down"
      Case WSAENETUNREACH
         strRet = "network is unreachable"
      Case WSAENETRESET
         strRet = "network dropped connection on reset"
      Case WSAECONNABORTED
         strRet = "software caused connection abort"
      Case WSAECONNRESET
         strRet = "connection reset by peer"
      Case WSAENOBUFS
         strRet = "no buffer space available"
      Case WSAEISCONN
         strRet = "socket is already connected"
      Case WSAENOTCONN
         strRet = "socket is not connected"
      Case WSAESHUTDOWN
         strRet = "cannot send after shutdown"
      Case WSAETOOMANYREFS
         strRet = "too many socket references"
      Case WSAETIMEDOUT
         strRet = "request timed out"
      Case WSAECONNREFUSED
         strRet = "connection refused"
      Case WSAENAMETOOLONG
         strRet = "name is too long"
      Case WSAEHOSTDOWN
         strRet = "host is down"
      Case WSAEHOSTUNREACH
         strRet = "host is unreachable"
      Case WSAEPROCLIM
         strRet = "too many processes"
      Case WSASYSNOTREADY
         strRet = "network sub-system is unavailable"
      Case WSAVERNOTSUPPORTED
         strRet = "requested version not supported"
      Case WSANOTINITIALISED
         strRet = "winsock is not loaded - call WSAStartup"
      Case WSAHOST_NOT_FOUND
         strRet = "host not found"
      '
      Case Else
         strRet = "unknown error"
      '
   End Select
   '
   vbGetLastError = strRet
   '
End Function
'
Public Function vbInetAddr(ByVal strIPAddress As String) As Long
   '
   ' Convert a dotted IP address into a network byte integer.
   vbInetAddr = inet_addr(strIPAddress)
   '
End Function
'
Public Function vbInetNtoa(ByVal lngIPAddress As Long) As String
   '
Dim lpString   As Long
Dim strBuffer  As String
   '
   ' Return a dotted 4 octet address from a 32bit network byte integer.
   lpString = inet_ntoa(lngIPAddress)
   If (lpString) Then
      '
      ' Prepare a buffer, copy the IP into it, then trim and return.
      strBuffer = String$(16, 0)
      Call CopyMemory(ByVal strBuffer, ByVal lpString, Len(strBuffer))
      vbInetNtoa = Mid$(strBuffer, 1, InStr(1, strBuffer, Chr$(0)) - 1)
      '
   End If
   '
End Function
'
Public Function vbHostNameFromIP(ByVal strIPAddress As String) As String
   '
Dim udtHost       As hostent
Dim lngIPAddress  As Long
Dim lngPointer    As Long
Dim strBuffer     As String
   '
   ' Resolve a dotted IP address into a hostname.
   '
   ' First, convert the string IP to a long IP.
   lngIPAddress = vbInetAddr(strIPAddress)
   If (lngIPAddress = INADDR_NONE) Then Exit Function
   '
   ' Now call gethostbyaddr to retrieve the hostent structure.
   lngPointer = gethostbyaddr(lngIPAddress, 4, AF_INET)
   If (lngPointer) Then
      '
      ' Copy the hostent structure out of the pointer.
      Call CopyMemory(udtHost, ByVal lngPointer, LenB(udtHost))
      '
      ' Prepare a string buffer and copy the hostname into it from the
      ' hostent.h_name field.
      strBuffer = String$(1024, 0)
      Call CopyMemory(ByVal strBuffer, ByVal udtHost.h_name, Len(strBuffer))
      '
      ' Trim the null characters off, and return the buffer.
      vbHostNameFromIP = Mid$(strBuffer, 1, InStr(1, strBuffer, Chr$(0)) - 1)
      '
   End If
   '
End Function
'
Public Function vbIPFromHostName(ByVal strHostName As String) As String
   '
Dim udtHost                As hostent
Dim lngIPAddress           As Long
Dim lngPointer             As Long
Dim bytIPAddress(0 To 3)   As Byte
Dim strBuffer              As String
Dim i                      As Long
   '
   ' Resolve a hostname into a dotted IP address.
   '
   ' Firstly, check if the hostname is already an IP.
   lngIPAddress = vbInetAddr(strHostName)
   If (lngIPAddress <> INADDR_NONE) Then
      '
      ' If it's already an IP, just return it.
      vbIPFromHostName = strHostName
      Exit Function
      '
   End If
   '
   ' It's not an IP, so we'll have to resolve it. Call gethostbyname().
   lngPointer = gethostbyname(strHostName)
   If (lngPointer) Then
      '
      ' Copy the hostent structure to local memory.
      Call CopyMemory(udtHost, ByVal lngPointer, LenB(udtHost))
      '
      ' h_addr_list contains a pointer to a long. So, firstly, copy out the
      ' pointer.
      Call CopyMemory(lngPointer, ByVal udtHost.h_addr_list, udtHost.h_length)
      '
      ' Copy the IP address into a four byte array, so we can build a
      ' dotted IP string from it.
      Call CopyMemory(bytIPAddress(0), ByVal lngPointer, udtHost.h_length)
      '
      ' Build and return the IP string.
      For i = 0 To 3
         strBuffer = strBuffer & CStr(bytIPAddress(i)) & "."
      Next i
      vbIPFromHostName = Mid$(strBuffer, 1, Len(strBuffer) - 1)
      '
   End If
   '
End Function
'
Public Function vbIsHostAlive(ByVal strHostAddress As String, _
                              ByVal lngWaitMilliseconds As Long) As Long
   '
Dim hEcho            As Long
Dim strIPAddress     As String
Dim lngIPAddress     As Long
Dim udtEchoRequest   As IP_OPTION_INFORMATION
Dim udtEchoReply     As IP_ECHO_REPLY
   '
   ' Ping the host to see if it's alive. Return the time.
   '
   ' Create an ICMP echo handle.
   hEcho = IcmpCreateFile()
   If (hEcho) Then
      '
      ' Convert the hostname (or IP address) into a long IP.
      strIPAddress = vbIPFromHostName(strHostAddress)
      lngIPAddress = vbInetAddr(strIPAddress)
      If (lngIPAddress <> INADDR_NONE) Then
         '
         ' Setup the echo options header.
         udtEchoRequest.TTL = 255
         '
         ' Send the echo.
         Call IcmpSendEcho(hEcho, _
                           lngIPAddress, _
                           vbNullString, _
                           0, _
                           udtEchoRequest, _
                           udtEchoReply, _
                           LenB(udtEchoReply), _
                           lngWaitMilliseconds)
         '
         ' Return the time it took. If the host is not alive, this will be 0.
         vbIsHostAlive = udtEchoReply.RoundTripTime
         '
      End If
      '
      ' Release the ICMP echo resources.
      Call IcmpCloseHandle(hEcho)
      '
   End If
   '
End Function
'


