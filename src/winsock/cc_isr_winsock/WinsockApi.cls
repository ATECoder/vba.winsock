VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinsockApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <summary>   Creates a socket that is bound to a specific transport service provider. </summary>
''' <param name="addressFamily">   [in] The address family specification.
''' The values currently supported are <see cref="ws32_AF_INET"/>   or <see cref="ws32_AF_INET6"/>, which are the Internet
''' address family formats for IPv4 and IPv6. Other options for address family (ws32_AF_NETBIOS for use with NetBIOS,
''' for example) are supported if a Windows Sockets service provider for the address family is installed.
''' Note that the values for the ws32_AF_ address family and PF_ protocol family constants are identical
''' (for example, ws32_AF_INET and PF_INET), so either constant can be used.
''' </param>
''' <param name="socketType">   [in] The type specification for the new socket.
''' In Windows Sockets 1.1, the only possible socket types are ws32_SOCK_DGRAM and ws32_SOCK_STREAM. </param>
''' <param name="protocol">   The protocol to be used. The possible options for the protocol parameter are specific
''' to the address family and socket type specified. </param>
''' <returns>   If no error occurs, socket returns a descriptor referencing the new socket.
''' Otherwise, a value of ws32_INVALID_SOCKET is returned, and a specific error code can be retrieved by
''' calling WSAGetLastError. <returns>
Private Declare PtrSafe Function CreateSocket_ Lib "wsock32.dll" Alias "socket" ( _
        ByVal a_addressFamily As Long, ByVal a_socketType As Long, ByVal a_protocol As Long) As Long

''' <summary>   Converts a u_short from host to TCP/IP network byte order (which is big-endian). </summary>
''' <remarks>
''' The htons function takes a 16-bit number in host byte order and returns a 16-bit number in network byte order
''' used in TCP/IP networks (the ws32_AF_INET or ws32_AF_INET6 address family).
'''
''' The htons function can be used to convert an IP port number in host byte order to the IP port number
''' in network byte order.
'''
''' The htons function does not require that the Winsock DLL has previously been loaded with a successful call
''' to the WSAStartup function.
''' </remarks>
''' <param name="a_host">   [in] A 16-bit unsigned number (represented as long) in host byte order. </param>
''' <returns>   the value in TCP/IP network byte order. <returns>
Private Declare PtrSafe Function ToInetByteOrder_ Lib "wsock32.dll" Alias "htons" ( _
        ByVal a_host As Long) As Integer
        
''' <summary>   Converts an IPv4 dotted-decimal address into a binary address value for the ws32_IN_ADDR structure. </summary>
''' <param name="a_hostName">   [in] An IPv4 dotted-decimal address. </param>
''' <returns>   If no error occurs, returns an unsigned long value containing a suitable binary
''' representation of the Internet address given. If the string in the hostname parameter does not contain a legitimate
''' Internet address, for example if a portion of an "a.b.c.d" address exceeds 255, then returns the value
''' ws32_INADDR_NONE. <returns>
Private Declare PtrSafe Function ToInetAddress_ Lib "wsock32.dll" Alias "inet_addr" ( _
        ByVal a_hostName As String) As Long
        
''' <summary>   Determines the status of one or more sockets, waiting if necessary, to perform synchronous I/O. </summary>
''' <param name="a_NFDS">       [in] Ignored. The nfds parameter is included only for compatibility with Berkeley sockets.</param>
''' <param name="a_readFSD">    [in, out] An optional pointer to a set of sockets to be checked for readability. </param>
''' <param name="a_writeFDS">   [in, out] An optional pointer to a set of sockets to be checked for writability. </param>
''' <param name="a_errorFDS">   [in, out] An optional pointer to a set of sockets to be checked for errors. </param>
''' <param name="a_timeout">    [in] The maximum time for select to wait, provided in the form of a TIMEVAL structure.
'''                             Set the timeout parameter to null for blocking operations. </param>
''' <returns>
''' The total number of socket handles that are ready and contained in the ws32_fd_set structures, zero if the time limit expired,
''' or ws32_SOCKET_ERROR if an error occurred. If the return value is ws32_SOCKET_ERROR, WSAGetLastError can be used to retrieve
''' a specific error code.
''' <returns>
Private Declare PtrSafe Function Select_ Lib "wsock32.dll" Alias "select" ( _
        ByVal a_nfds As Integer, ByRef a_readFDS As wsock32.ws32_fd_set, ByRef a_writeFDS As wsock32.ws32_fd_set, _
        ByRef a_errorFDS As wsock32.ws32_fd_set, ByRef a_timeout As wsock32.ws32_Time_Value) As Integer

''' <summary>   Creates a socket that is bound to a IPv4 transport service. </summary>
''' <returns>   If no error occurs, socket returns a descriptor referencing the new socket.
''' Otherwise, a value of ws32_INVALID_SOCKET is returned, and a specific error code can be retrieved by
''' calling WSAGetLastError. <returns>
Public Function CreateIPv4StreamSocket() As Long
    CreateIPv4StreamSocket = CreateSocket_(wsock32.ws32_AF_INET, wsock32.ws32_SOCK_STREAM, 0)
End Function

''' <summary>   Converts an unsigned short value from host to TCP/IP network byte order (which is big-endian). </summary>
''' <param name="a_value">   [in] A 16-bit number in host byte order. </param>
''' <returns>   The value in TCP/IP network byte order. <returns>
Public Function ToInetByteOrder(ByVal a_value As Long) As Integer
    ToInetByteOrder = ToInetByteOrder_(a_value)
End Function

''' <summary>   Converts an IPv4 dotted-decimal address into a binary address value for the ws32_IN_ADDR structure. </summary>
''' <param name="a_host">   [in] An IPv4 dotted-decimal address. </param>
''' <returns>   If no error occurs, returns an unsigned long value containing a suitable binary
''' representation of the Internet address given. If the string in the hostname parameter does not contain a legitimate
''' Internet address, for example if a portion of an "a.b.c.d" address exceeds 255, then returns the value
''' ws32_INADDR_NONE. <returns>
Public Function ToInetAddress(ByVal a_host As String) As Long
    ToInetAddress = ToInetAddress_(a_host)
End Function

''' <summary>   Initializes the socket set to the empty set. A set should always be cleared before using./ </summary>
''' <param name="socketSet">   [in, out] A pointer to the set of sockets. </param>
Public Sub FD_SET_INIT(ByRef a_socketSet As wsock32.ws32_fd_set)
    a_socketSet.fd_count = 0
End Sub

''' <summary>   Removes socket s from set. </summary>
''' <param name="a_socket_Id">   [in] A descriptor that identifies a socket. </param>
''' <param name="a_socketSet">   [in, out] A pointer to the set of sockets. </param>
Public Sub FD_SET_REMOVE(ByVal a_socket_id As Long, ByRef a_socketSet As wsock32.ws32_fd_set)
    
    Dim i As Integer
    i = 0
    Do While i < a_socketSet.fd_count
        If a_socketSet.fd_array(i) = a_socket_id Then
            a_socketSet.fd_array(i) = 0
            a_socketSet.fd_count = a_socketSet.fd_count - 1
            Exit Do
        End If
        i = i + 1
    Loop

End Sub

''' <summary>   Checks to see if s is a member of set and returns TRUE if so.. </summary>
''' <param name="socket_Id">   [in] A descriptor that identifies a socket. </param>
''' <param name="socketSet">   [in, out] A pointer to the set of sockets. </param>
Public Function FD_SET_CONTAINS(ByVal a_socket_id As Long, ByRef a_socketSet As wsock32.ws32_fd_set) As Boolean
    
    Dim p_found As Boolean: p_found = False
    
    Dim i As Integer
    i = 0
    
    Do While i < a_socketSet.fd_count
        If a_socketSet.fd_array(i) = a_socket_id Then
            p_found = True
            Exit Do
        End If
        i = i + 1
    Loop
    FD_SET_CONTAINS = p_found
   
End Function

''' <summary>   Adds a socket to the set. </summary>
''' <param name="a_socket_Id">   [in] A descriptor that identifies a socket. </param>
''' <param name="a_socketSet">   [in, out] A pointer to the set of sockets. </param>
Public Sub FD_SET_ADD(ByVal a_socket_id As Long, ByRef a_socketSet As wsock32.ws32_fd_set)
    
    Dim i As Integer: i = 0
    
    Do While i < a_socketSet.fd_count
        If a_socketSet.fd_array(i) = a_socket_id Then
            Exit Do
        End If
        i = i + 1
    Loop
    
    If i = a_socketSet.fd_count Then
        If a_socketSet.fd_count < wsock32.ws32_FD_SETSIZE Then
            a_socketSet.fd_array(i) = a_socket_id
            a_socketSet.fd_count = a_socketSet.fd_count + 1
        End If
    End If
    
End Sub

''' <summary>   Determines the readability status of one or more sockets, waiting if necessary, to perform synchronous I/O. </summary>
''' <param name="readSocketSet">   [in, out] A pointer to a set of sockets to be checked for readability. </param>
''' <param name="timeoutMs">       [Optional, 500] [in] The maximum milliseconds time for select to wait. </param>
''' <returns>
''' The total number of socket handles that are ready and contained in the ws32_fd_set structures, zero if the time limit expired,
''' or ws32_SOCKET_ERROR if an error occurred. If the return value is ws32_SOCKET_ERROR, WSAGetLastError can be used to retrieve
''' a specific error code.
''' <returns>
Public Function DetermineReadability(ByRef a_readSocketSet As wsock32.ws32_fd_set, _
                                     Optional ByVal a_timeoutMs As Long = 500) As Integer

    Dim p_time As wsock32.ws32_Time_Value
    p_time.tv_sec = a_timeoutMs / 1000
    p_time.tv_usec = 1000 * (a_timeoutMs Mod 1000)
    
    Dim p_emptyFdSet As wsock32.ws32_fd_set
    
    Dim p_selectResult As Integer
    p_selectResult = Select_(0, a_readSocketSet, p_emptyFdSet, p_emptyFdSet, p_time)

    If p_selectResult = wsock32.ws32_SOCKET_ERROR Then
        Err.Raise cc_isr_Core.UserDefinedErrors.SocketSetStatusError.Code, "IdentityReader.WinsockApi.DetermineReadability", _
                  cc_isr_Core.UserDefinedErrors.SocketSetStatusError.Description & _
                  Winsock.LastErrorMessage(" Last error: ")
    End If
    DetermineReadability = p_selectResult

End Function

