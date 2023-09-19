Attribute VB_Name = "wsock32"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Winsock 32 constants. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   The Winsock implementation version </summary.
''' <remarks>
''' Version 1.1 (1*256 + 1) = 257
''' version 2.0 (2*256 + 0) = 512
''' </remarks>
Public Const ws32_VERSION = 257

Public Const ws32_WSADESCRIPTION_LEN = 256
Public Const ws32_WSASYS_STATUS_LEN = 128

Public Const ws32_WSADESCRIPTION_LEN_ARRAY = ws32_WSADESCRIPTION_LEN + 1
Public Const ws32_WSASYS_STATUS_LEN_ARRAY = ws32_WSASYS_STATUS_LEN + 1

''' <summary>   A data structure that receives the information returned from
''' the WSAStartup() function. </summary>
Public Type ws32_WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * ws32_WSADESCRIPTION_LEN_ARRAY
    szSystemStatus As String * ws32_WSASYS_STATUS_LEN_ARRAY
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type

' Define address families
Public Const ws32_AF_UNSPEC = 0             ' unspecified
Public Const ws32_AF_UNIX = 1               ' local to host (pipes, portals)
Public Const ws32_AF_INET = 2               ' The Internet Protocol version 4 (IPv4) address family.

' Define socket types

''' <summary>   A socket type that provides sequenced, reliable, two-way, connection-based byte streams with an
''' OOB data transmission mechanism. This socket type uses the Transmission Control Protocol (TCP) for the
''' Internet address family (ws32_AF_INET or ws32_AF_INET6). </summary>
Public Const ws32_SOCK_STREAM = 1

''' <summary>
''' A socket type that supports datagrams, which are connectionless, unreliable buffers of a fixed (typically
''' small) maximum length. This socket type uses the User Datagram Protocol (UDP) for the Internet address family
''' (ws32_AF_INET or ws32_AF_INET6).
''' </summary>
Public Const ws32_SOCK_DGRAM = 2

Public Const ws32_SOCK_RAW = 3              ' Raw data socket
Public Const ws32_SOCK_RDM = 4              ' Reliable Delivery socket
Public Const ws32_SOCK_SEQPACKET = 5        ' Sequenced Packet socket

Public Const ws32_INADDR_ANY As Long = 0
Public Const ws32_INADDR_NONE As Long = &HFFFFFFFF

''' <summary>   Sets the Internet address type as a long integer (32-bit) </summary>
Public Type ws32_IN_ADDR
    s_addr As Long
End Type

''' <summary>   Sets the socket IPv4 address expressed in network byte order. </summary>
Public Type ws32_Address
    sa_family As Integer
    sa_data As String * 14
End Type

Public Const ws32_AddressLen = 16

''' <summary>   Sets the socket IPv4 address expressed in network byte order. </summary>
Public Type ws32_Address_in
    sin_family As Integer     ' Address family of the socket, such as ws32_AF_INET.
    sin_port As Integer       ' sock address port number, e.g.,  htons(5150);
    sin_addr As ws32_IN_ADDR  ' the internet address as a long integer type.
    sin_zero As String * 8
End Type

Public Type ws32_Time_Value
    tv_sec As Long
    tv_usec As Long
End Type

' Define socket return codes
Public Const ws32_INVALID_SOCKET = &HFFFF
Public Const ws32_SOCKET_ERROR = -1

Public Const ws32_SOL_SOCKET = 65535   ' socket options
Public Const ws32_SO_RCVTIMEO = &H1006 ' receive timeout option

Public Const ws32_MSG_OOB = &H1       ' Process out-of-band data.
Public Const ws32_MSG_PEEK = &H2      ' Peek at incoming messages.
Public Const ws32_MSG_DONTROUTE = &H4 ' Don't use local routing
Public Const ws32_MSG_WAITALL = &H8   ' do not complete until packet is completely filled

Public Const ws32_FD_SETSIZE = 64

''' <summary>   Defines a set of sokkets returned from the <see cref="SelectSockets"/>   alias select
''' Winsock API call. </summary>
''' <remarks>
''' Four macros are defined in the header file Winsock2.h for manipulating and checking the descriptor sets.
''' The variable ws32_FD_SETSIZE determines the maximum number of descriptors in a set.
''' (The default value of ws32_FD_SETSIZE is 64, which can be modified by defining ws32_FD_SETSIZE
''' to another value before including Winsock2.h.)
''' Internally, socket handles in an ws32_fd_set structure are not represented as bit flags as in Berkeley Unix.
''' Their data representation is opaque. Use of these macros will maintain software portability between
''' different socket environments. The macros to manipulate and check ws32_fd_set contents are:
''' FD_ZERO(*set)     - (FD_SET_INIT) - Initializes set to the empty set. A set should always be cleared before using.
''' FD_CLR(s, *set)   - (FD_SET_REMOVE) - Removes socket s from set.
''' FD_ISSET(s, *set) - (FD_SET_CONTAINS) - Checks to see if s is a member of set and returns TRUE if so.
''' FD_SET(s, *set)   - (FD_SET_ADD) - Adds socket s to set.
''' </remarks>
Public Type ws32_fd_set
    fd_count As Integer
    fd_array(ws32_FD_SETSIZE) As Long
End Type

