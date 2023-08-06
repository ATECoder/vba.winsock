VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPv4StreamSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SEND_MESSAGE_ERROR = -1
Private Const SEND_PARTIAL_MESSAGE_ERROR = -2
Private Const RECEIVE_ERROR = -1
Private Const DEFAULT_TIMEOUT_MS = 500

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   Permits an incoming connection attempt on a socket. </summary>
''' <param name="a_socketId">             [in] A descriptor that identifies a socket that has been placed in a listening state with the listen function.
'''                                       The connection is actually made with the socket that is returned by accept. </param>
''' <param name="a_clientAddress">        [out] An optional pointer to a buffer that receives the address of the connecting entity,
'''                                        as known to the communications layer. The exact format of the address parameter is determined by the address family
'''                                        that was established when the socket from the ws32_SocketAddress structure was created. </param>
''' <param name="a_clientAddressLength">   [in, out] An optional pointer to an integer that contains the length of structure pointed to by
'''                                        the address parameter. </param>
''' <returns>   If no error occurs, accept returns a value of type SOCKET that is a descriptor for the new socket.
''' This returned value is a handle for the socket on which the actual connection is made. Otherwise, a value of
''' ws32_INVALID_SOCKET is returned, and a specific error code can be retrieved by calling WSAGetLastError.
''' The integer referred to by clientAddressLength initially contains the amount of space pointed to by clientAddress. On return it
''' will contain the actual length in bytes of the address returned. <returns>
Private Declare PtrSafe Function accept_ Lib "wsock32.dll" Alias "accept" ( _
        ByVal a_socketId As Long, ByRef a_clientAddress As wsock32.ws32_SocketAddress, _
        ByVal a_clientAddressLength As Integer) As Long

''' <summary>   Associates a local address with a socket. </summary>
''' <param name="a_socketid">        [in] A descriptor identifying an unbound socket. </param>
''' <param name="a_address">         [in] A pointer to a ws32_SocketAddress_in structure of the local address
'''                                  to assign to the bound socket . </param>
''' <param name="a_addressLength">   [in] The length, in bytes, of the value pointed to by address. </param>
''' <returns>   Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function bind_ Lib "wsock32.dll" Alias "bind" ( _
        ByVal a_socketId As Long, ByRef a_address As wsock32.ws32_SocketAddress_in, _
        ByVal a_addressLength As Integer) As Long

''' <summary>   Closes an existing socket. </summary>
''' <remarks>
''' The closesocket function closes a socket. Use it to release the socket descriptor passed in the
''' s parameter. Note that the socket descriptor passed in the s parameter may immediately be reused
''' by the system as soon as closesocket function is issued. As a result, it is not reliable to expect
''' further references to the socket descriptor passed in the s parameter to fail with the error WSAENOTSOCK.
''' A Winsock client must never issue closesocket on s concurrently with another Winsock function call.
''' </remarks>
''' <param name="a_socketId">   [in] A descriptor identifying the socket to close. </param>
''' <returns>   If no error occurs, closesocket returns zero. Otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function closesocket_ Lib "wsock32.dll" Alias "closesocket" ( _
        ByVal a_socketId As Long) As Long

''' <summary>   Establishes a connection to a specified socket. </summary>
''' <param name="a_socketId">     [in] A descriptor identifying an unconnected socket. </param>
''' <param name="a_address">      [in] A pointer to the <see cref="ws32_SocketAddress_in"/>   structure to which the
'''                               connection should be established. </param>
''' <param name="a_addressLen">   [in] The length, in bytes, of the ws32_SocketAddress structure pointed to by the
'''                               <paramref name="address"/>   parameter. </param>
''' <returns>   Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function connect_ Lib "wsock32.dll" Alias "connect" ( _
        ByVal a_socketId As Long, ByRef a_address As wsock32.ws32_SocketAddress_in, _
        ByVal a_addressLen As Long) As Long

''' <summary>   Places a socket in a state in which it is listening for an incoming connection. </summary>
''' <param name="a_socketId">   [in] A descriptor identifying a bound, unconnected socket. </param>
''' <param name="a_backlog">    [in] The maximum length of the queue of pending connections. If set to SOMAXCONN,
'''                             the underlying service provider responsible for socket s will set the backlog to a
'''                             maximum reasonable value. If set to SOMAXCONN_HINT(N) (where N is a number), the
'''                             backlog value will be N, adjusted to be within the range (200, 65535). Note that
'''                             SOMAXCONN_HINT can be used to set the backlog to a larger value than possible with SOMAXCONN.
'''                             SOMAXCONN_HINT is only supported by the Microsoft TCP/IP service provider. There is no
'''                             standard provision to obtain the actual backlog value.
''' </param>
''' <returns>   Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function listen_ Lib "wsock32.dll" Alias "listen" ( _
        ByVal a_socketId As Long, ByVal a_backlog As Integer) As Long

''' <summary>   Sends data on a connected socket. </summary>
''' <param name="a_socketId">       [in] A descriptor identifying a connected socket. </param>
''' <param name="a_buffer">         [in] A pointer to a buffer containing the data to be transmitted. </param>
''' <param name="a_bufferLength">   [in] The length, in bytes, of the data in buffer pointed to by the buffer parameter. </param>
''' <param name="a_flags">          [in] A set of flags that specify the way in which the call is made. This parameter is
'''                                 constructed by using the bitwise OR operator with any of the following values:
'''                                 ws32_MSG_DONTROUTE: Specifies that the data should not be subject to routing. A Windows Sockets
'''                                 service provider can choose to ignore this flag.
'''                                 ws32_MSG_OOB: Sends OOB data (stream-style socket such as ws32_SOCK_STREAM only).
'''                                 </param>
''' <returns>   Returns the total number of bytes sent, which can be less than the number requested to be sent
''' in the length parameter. If an error occurred, a value of ws32_SOCKET_ERROR is returned, and a specific error code
''' can be retrieved by calling WSAGetLastError.
''' <returns>
Private Declare PtrSafe Function send_ Lib "wsock32.dll" Alias "send" ( _
        ByVal a_socketId As Long, ByRef a_buffer As String, ByVal a_bufferLength As Long, _
        ByVal a_flags As Long) As Long

''' <summary>   Sets a socket option. </summary>
''' <param name="a_socketId">       [in] A descriptor that identifies a socket. </param>
''' <param name="a_level">          [in] The level at which the option is defined (for example, ws32_SOL_SOCKET). </param>
''' <param name="a_opttionName">    [in] The socket option for which the value is to be set (for example, ws32_SO_BROADCAST).
'''                                 The option name parameter must be a socket option defined within the specified level,
'''                                 or behavior is undefined. </param>
''' <param name="a_optiobValue">    [in] A pointer to the buffer in which the value for the requested option is specified. </param>
''' <param name="a_optionLength">   [in] The size, in bytes, of the buffer pointed to by the optval parameter. </param>
''' <returns>   If no error occurs, setsockopt returns zero. Otherwise, a value of ws32_SOCKET_ERROR is returned,
''' and a specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function setsockopt_ Lib "wsock32.dll" Alias "setsockopt" ( _
        ByVal a_socketId As Long, ByVal a_level As Long, ByVal a_optionName As Long, _
        ByRef a_optionValue As Long, ByVal a_optLength As Integer) As Long

''' <summary>   Receives data from a connected socket or a bound connectionless socket. </summary>
''' <remarks>
''' The flags parameter can be used to influence the behavior of the function invocation beyond the options specified
''' for the associated socket. The semantics of this function are determined by the socket options and the flags parameter.
''' The possible value of flags parameter is constructed by using the bitwise OR operator with any of the following values:
''' ws32_MSG_PEEK    Peeks at the incoming data. The data is copied into the buffer, but is not removed from the input queue.
''' ws32_MSG_OOB     Processes Out Of Band (OOB) data.
''' ws32_MSG_WAITALL The receive request will complete only when one of the following events occurs:
'''                  The buffer supplied by the caller is completely full.
'''                  The connection has been closed.
'''                  The request has been canceled or an error occurred.
''' Note that if the underlying transport does not support ws32_MSG_WAITALL, or if the socket is in a non-blocking mode,
''' then this call will fail with WSAEOPNOTSUPP. Also, if ws32_MSG_WAITALL is specified along with ws32_MSG_OOB, ws32_MSG_PEEK, or
''' ws32_MSG_PARTIAL, then this call will fail with WSAEOPNOTSUPP. This flag is not supported on datagram sockets
''' or message-oriented sockets.
''' </remarks>
''' <param name="a_socketId">       [in] A descriptor identifying a connected socket. </param>
''' <param name="a_buffer">         [out] A pointer to the buffer to receive the incoming data. </param>
''' <param name="a_bufferLength">   [in] The length, in bytes, of the data in buffer pointed to by the buffer parameter. </param>
''' <param name="a_flags">          [in] A set of flags that influences the behavior of this function. </param>
''' <returns>
''' If no error occurs, receive returns the number of bytes received and the buffer pointed to by the buffer parameter will
''' contain this data received. If the connection has been gracefully closed, the return value is zero.
''' Otherwise, a value of ws32_SOCKET_ERROR is returned, and a specific error code can be retrieved by calling WSAGetLastError.
''' <returns>
Private Declare PtrSafe Function recv_ Lib "wsock32.dll" Alias "recv" ( _
        ByVal a_socketId As Long, ByVal a_buffer As String, ByVal a_bufferLength As Long, _
        ByVal a_flags As Long) As Long

Private m_socketPort As Long
Private m_socketHost As String
Private m_socketAddress As String
Private m_socketId As Long
Private m_connected As Boolean
Private m_listening As Boolean
Private m_receiveTimeout As Long
Private m_fdSet As wsock32.ws32_fd_set

''' <summary>   Initializes an internet protocol Stream Socket with a socket id. </summary>
Private Sub Class_Initialize()

    ' see that the singleton Winsock class is instantiated
    On Error Resume Next
    
    ' this is required to initialize Winsock.  It will only ran once.
    Winsock.Initialize
    
    m_socketHost = vbNullString
    m_socketAddress = vbNullString
    
    On Error GoTo 0

    ' Create a new socket
    
    m_socketId = WinsockApi.CreateIPv4StreamSocket()
    If m_socketId = wsock32.ws32_INVALID_SOCKET Then
        Err.Raise cc_isr_Core.UserDefinedErrors.SocketCreationError.Code, "IdentityReader.IPv4StreamSocket.Initialize", _
                  cc_isr_Core.UserDefinedErrors.SocketCreationError.Description & _
                  Winsock.LastErrorMessage(" Last error: ")
    Else
        Winsock.Register Me
        
        ' set the receive timeout to its default value
        
        Me.SetReceiveTimeout DEFAULT_TIMEOUT_MS
    End If
    
End Sub

''' <summary>   terminates the socket and releases its Winsock registration. </summary>
Private Sub Class_Terminate()

    ' see that the singleton Winsock class is instantiated
    On Error Resume Next
    
    ' unregister this socket.
    Winsock.UnRegister Me
    
    On Error GoTo 0

End Sub

''' <summary>   Returns the port the server is listening to. </summary>
''' <value>   [Long] The port the server listens to. </value>
Public Property Get SocketPort() As Long
    SocketPort = m_socketPort
End Property

''' <summary>   Gets the IPv4 dotted-decimal host address. </summary>
''' <value>   [String] The IPv4 dotted-decimal host address. </value>
Public Property Get SocketHost() As String
    SocketHost = m_socketHost
End Property

''' <summary>   Gets the socket address consisting of the IPv4 dotted-decimal host
''' and the port the server is listening too. </summary>
''' <value>   [String] The IPv4 dotted-decimal host address and port delimited with a colon. </value>
Public Property Get SocketAddress() As String
    SocketAddress = m_socketAddress
End Property

''' <summary>   Returns a descriptor identifying the socket. </summary>
''' <value>   The socket identifying descriptor. </value>
Public Property Get SocketId() As Long
Attribute SocketId.VB_Description = "A descriptor identifying the socket"
Attribute SocketId.VB_UserMemId = 0
    SocketId = m_socketId
End Property

''' <summary>   Returns the connection state of the socket. </summary>
''' <value>   True if the socket is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
    Connected = m_connected
End Property

''' <summary>   Initializes an Inet stream socket using the provided socket id. </summary>
''' <param name="connectedSocketId">   The ID of a connected socket. </param>
''' <param name="timeoutMs">           [Optional, 500] The receive timeout interval in milliseconds. </param>
Public Sub Initialize(ByVal a_connectedSocketId As Long, Optional ByVal a_timeoutMs As Long = 500)
    
    m_socketId = a_connectedSocketId
    m_connected = True
    Me.SetReceiveTimeout a_timeoutMs

End Sub

''' <summary>   Opens a socket connection and sets and returns the socket id. </summary>
''' <param name="a_host">   [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">   [Long] The port that the server is listening on. </param>
''' <returns>   [Boolean] True if the socket is connected. </returns>
Public Function OpenConnection(ByVal a_host As String, ByVal a_port As Long) As Boolean
   
    m_connected = False
    If StringExtensions.IsNullOrEmpty(a_host) Then
        Err.Raise cc_isr_Core.UserDefinedErrors.InvalidArgumentError.Code, "IdentityReader.IPv4StreamSocket.OpenConnection", _
                  cc_isr_Core.UserDefinedErrors.InvalidArgumentError.Description & " Argument: a_host"
    End If
    
    If a_host = "localhost" Then: a_host = "127.0.0.1"
    
    ' Open a connection to a server
    ' can use this: https://www.allscoop.com/Tcp-listen.php to listen.
    
    Dim p_address As wsock32.ws32_SocketAddress_in
    
    ' set the binary address
    p_address.sin_addr.s_addr = WinsockApi.ToInetAddress(a_host)
    
    p_address.sin_family = wsock32.ws32_AF_INET
    
    ' set the port in network byte order
    p_address.sin_port = WinsockApi.ToInetByteOrder(a_port)
    
    Dim p_connectResult As Long
    p_connectResult = connect_(m_socketId, p_address, Len(p_address))
    
    If p_connectResult = wsock32.ws32_SOCKET_ERROR Then
        m_connected = False
        Err.Raise cc_isr_Core.UserDefinedErrors.SocketConnectionError.Code, "IdentityReader.IPv4StreamSocket.OpenConnection", _
                  cc_isr_Core.UserDefinedErrors.SocketConnectionError.Description & _
                  Winsock.LastErrorMessage(" Last error: ")
    Else
        m_connected = True
        
        OnConnectionChanged Me.Connected

    End If
    
    m_socketPort = a_port
    m_socketHost = a_host
    m_socketAddress = a_host & ":" & CStr(a_port)
    
    OpenConnection = m_connected

End Function

''' <summary>   Closes the socket connection and zeros the socket id. </summary>
''' <returns>   [Boolean] True if socket disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean

    Dim p_result As Long
    p_result = closesocket_(m_socketId)
    
    ' set the socket id to invalid to indicate that it is no longer viable.
    m_socketId = wsock32.ws32_INVALID_SOCKET
    
    If p_result = wsock32.ws32_SOCKET_ERROR Then
        CloseConnection = False
        Err.Raise cc_isr_Core.UserDefinedErrors.SocketDisconnectionError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                  cc_isr_Core.UserDefinedErrors.SocketDisconnectionError.Description & _
                  Winsock.LastErrorMessage(" Last error: ")
    Else
        m_connected = False
        
        CloseConnection = True
        
        OnConnectionChanged Me.Connected
        
    End If

End Function

''' <summary>   Associates a local address with a socket. </summary>
''' <param name="a_address">   The INet address to bind to. </param>
''' <param name="a_port">   The port that the server is listening on. </param>
''' <returns>   [Boolean] True if the socket is bound. </returns>
Public Function BindTo(ByVal a_address As Long, ByVal a_port As Long) As Boolean

    Dim p_endpoint As wsock32.ws32_SocketAddress_in
    p_endpoint.sin_family = wsock32.ws32_AF_INET
    p_endpoint.sin_addr.s_addr = a_address
    p_endpoint.sin_port = WinsockApi.ToInetByteOrder(a_port)
    
    Dim p_bindResult As Long
    p_bindResult = bind_(Me.SocketId, p_endpoint, 16)
    
    If p_bindResult = wsock32.ws32_SOCKET_ERROR Then
        Err.Raise cc_isr_Core.UserDefinedErrors.SocketBindError.Code, "IdentityReader.IPv4StreamSocket.BindTo", _
                  cc_isr_Core.UserDefinedErrors.SocketBindError.Description & _
                  Winsock.LastErrorMessage(" Last error: ")
        m_connected = False
    Else
        m_connected = True
    End If
    BindTo = Me.Connected
   
End Function

''' <summary>   Returns the listening state of the socket. </summary>
''' <value>   [Boolean] True if the socket is Listening; otherwise, False. </value>
Public Property Get Listening() As Boolean
    Listening = m_listening
End Property


''' <summary>   Places a socket in a state in which it is listening for an incoming connection. </summary>
''' <param name="a_backlog">   [Optional, 10] [in] The maximum length of the queue of pending connections. </para>
''' <returns>   [Boolean] True if the socket is Listening; otherwise, False. </returns>
Public Function StartListening(Optional ByVal a_backlog As Integer = 10) As Boolean

    Dim p_listenResult As Long
    p_listenResult = listen_(Me.SocketId, a_backlog)
    
    If p_listenResult = wsock32.ws32_SOCKET_ERROR Then
        Err.Raise cc_isr_Core.UserDefinedErrors.SocketListenError.Code, "IdentityReader.IPv4StreamSocket.Listen", _
                  cc_isr_Core.UserDefinedErrors.SocketListenError.Description & _
                  Winsock.LastErrorMessage(" Last error: ")
        m_listening = False
    Else
        m_listening = True
    End If
    StartListening = Me.Listening

End Function

''' <summary>   Permits an incoming connection attempt on a socket. </summary>
''' <param name="a_serverSocket">   [<see cref="IPv4StreamSocket"/>] The server socket. </param>
''' <param name="a_timeoutMs">      [Optional, Long, 500] The maximum milliseconds time for select to wait. </param>
''' <returns>   the connected socket. </returns>
Public Function AcceptTcpClient(ByVal a_serverSocket As IPv4StreamSocket, _
        Optional ByVal a_timeoutMs As Long = 500) As IPv4StreamSocket

    ' add this socket to a socket set.
    WinsockApi.FD_SET_INIT m_fdSet
    WinsockApi.FD_SET_ADD m_socketId, m_fdSet
    
    ' check if we have available socket for readability
    
    Dim p_readableSocketCount As Integer
    p_readableSocketCount = WinsockApi.DetermineReadability(m_fdSet, a_timeoutMs)
    
    If p_readableSocketCount = 0 Then
        Set AcceptTcpClient = Nothing
        Exit Function
    End If
    
    Dim p_socketId As Long
    Dim p_socketAddress As wsock32.ws32_SocketAddress
    
    ' get a connected socket id
    
    p_socketId = accept_(a_serverSocket.SocketId, p_socketAddress, wsock32.ws32_SocketAddressLen)
    
    If p_socketId = wsock32.ws32_INVALID_SOCKET Then
        Err.Raise cc_isr_Core.UserDefinedErrors.SocketAcceptError.Code, "IdentityReader.IPv4StreamSocket.AcceptTcpClient", _
                  cc_isr_Core.UserDefinedErrors.SocketAcceptError.Description & _
                  Winsock.LastErrorMessage(" Last error: ")
        Exit Function
    End If
    
    Dim p_clientSocket As IPv4StreamSocket
    Set p_clientSocket = New IPv4StreamSocket
    p_clientSocket.Initialize p_socketId, a_timeoutMs
    
    Set AcceptTcpClient = p_clientSocket
    
End Function

''' <summary>   Sends a terminated message to the server. </summary>
''' <param name="a_message">   [String] The terminated message. </param>
''' <returns>   [Long} The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal a_message As String) As Long

    If StringExtensions.IsNullOrEmpty(a_message) Then
        SendMessage = 0
    Else
        Dim p_messageLength As Integer: p_messageLength = Len(a_message)
        Dim p_sentLength As Integer
        p_sentLength = send_(Me.SocketId, ByVal a_message, Len(a_message), 0)
        
        If p_sentLength < 0 Then
            ' if the sent length is negative, we have an API error.
            Err.Raise cc_isr_Core.UserDefinedErrors.SocketSendError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                      cc_isr_Core.UserDefinedErrors.SocketSendError.Description & _
                      Winsock.LastErrorMessage(" Last error: ")
            SendMessage = SEND_MESSAGE_ERROR
        ElseIf p_sentLength <> p_messageLength Then
            Err.Raise cc_isr_Core.UserDefinedErrors.SocketSendError.Code, "IdentityReader.IPv4StreamSocket.CloseConnection", _
                      cc_isr_Core.UserDefinedErrors.SocketSendError.Description & _
                      " A partial message was sent: " & CStr(p_sentLength) & " out of " & CStr(p_messageLength)
            SendMessage = SEND_PARTIAL_MESSAGE_ERROR
        Else
            SendMessage = p_sentLength
        End If
    End If

End Function

''' <summary>   Receives a message from the server until reaching
''' the specified termination. </summary>
''' <param name="a_buffer">        [String] Where the message will go. </param>
''' <param name="a_maxLength">     [Integer] The maximum number of bytes to read. </param>
''' <param name="a_termination">   [Optional, String, vbLf] The character that signifies the end of the stream. </param>
''' <returns>   [Integer] If no error, the number of byte that were sent to the server.
''' Otherwise, RECEIVE_ERROR.
''' </returns>
Public Function ReceiveTerminatedMessage(ByRef a_buffer As String, ByVal a_maxLength As Integer, _
                                         Optional ByVal a_termination As String = vbLf) As Integer

    Dim c As String * 1
    Dim p_length As Integer
    Dim p_count As Integer
    
    a_buffer = vbNullString
    While p_length < a_maxLength
    
        DoEvents
        
        c = vbNullString
        Dim l As Long: l = Len(c)
        p_count = recv_(Me.SocketId, c, l, 0)
        
        If p_count < 1 Then
            ReceiveTerminatedMessage = RECEIVE_ERROR
            a_buffer = Chr$(0)
            Exit Function
        End If
        
        If c = a_termination Then
           a_buffer = a_buffer + Chr$(0)
           ReceiveTerminatedMessage = p_length
           Exit Function
        End If
        
        p_length = p_length + p_count
        a_buffer = a_buffer + c
        
    Wend
    
    ReceiveTerminatedMessage = RECEIVE_ERROR
    
End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="a_maxLength">   [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">   [Optional, True] True to return the string without the termination. </param>
''' <param name="a_termination">   [Optional, 10 ASCII] The character that signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function ReceiveRaw(Optional ByVal a_maxLength As Long = &H7FFFF, _
                           Optional ByVal a_trimEnd As Boolean = True, _
                           Optional ByVal a_termination As String = vbLf) As String

    Dim c As String * 1
    Dim p_count As Integer
    
    Dim p_builder As StringBuilder: Set p_builder = cc_isr_Core.Constructor.CreateStringBuilder()
    Dim p_currentLength As Long: p_currentLength = 0
    
    Do While p_currentLength < a_maxLength
    
        DoEvents
        
        c = vbNullString
        p_count = recv_(Me.SocketId, c, 1, 0)
        
        If p_count < 1 Then
            Err.Raise cc_isr_Core.UserDefinedErrors.SocketReceiveError.Code, "IdentityReader.IPV4StreamSocket.receiveRaw", _
                      cc_isr_Core.UserDefinedErrors.SocketReceiveError.Description & _
                      Winsock.LastErrorMessage(" Last error: ")
            Exit Do
        ElseIf c = a_termination Then
            If Not a_trimEnd Then
                p_builder.Append c
            End If
            Exit Do
        Else
            p_currentLength = p_currentLength + p_count
            p_builder.Append c
        End If
    Loop
    ReceiveRaw = p_builder.ToString
    
End Function

''' <summary>   Receives a message from the server. </summary>
''' <param name="a_maxLength">   The maximum number of maxLength to read. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveBytes(ByVal a_maxLength As Long) As String
    
    Dim p_buffer As String: p_buffer = StringExtensions.Repeat(Chr(0), a_maxLength)
    
    Dim p_actualLength As Long
    
    p_actualLength = recv_(Me.SocketId, p_buffer, a_maxLength, 0)
    
    If p_actualLength > 0 Then
        ReceiveBytes = StringExtensions.Substring(p_buffer, 0, p_actualLength)
    Else
        ReceiveBytes = vbNullString
    End If
    
End Function

''' <summary>   Receives a message from the server. </summary>
''' <remarks>   Data is read until exhausted. </remarks>
''' <param name="a_bufferSize">   [Optional, 1024] The number of character to read on each trial. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveString(Optional ByVal a_bufferSize As Integer = 1024) As String
    
    Dim p_buffer As String
    Dim p_builder As cc_isr_Core.StringBuilder: p_builder = cc_isr_Core.Constructor.CreateStringBuilder
    
    Do
        p_buffer = Trim(Me.ReceiveBytes(a_bufferSize))
        If Len(p_buffer) > 0 Then p_builder.Append p_buffer
        
    Loop While Len(p_buffer) > 0
    
    ReceiveString = p_builder.ToString()
    
End Function

''' <summary>   Sets the socket receive timeout. </summary>
''' <param name="a_timeoutMs">   The timeout interval in milliseconds. </param>
Public Sub SetReceiveTimeout(ByVal a_timeoutMs As Long)

    Dim p_result As Long
    
    ' check if the requested receive timeout is different from the
    ' timeout that was already set.
    If m_receiveTimeout <> a_timeoutMs Then
    
        p_result = setsockopt_(m_socketId, wsock32.ws32_SOL_SOCKET, wsock32.ws32_SO_RCVTIMEO, a_timeoutMs, 4)
    
        If p_result = wsock32.ws32_SOCKET_ERROR Then
            Err.Raise cc_isr_Core.UserDefinedErrors.SocketOptionSettingError.Code, "IdentityReader.IPV4StreamSocket.SetReceiveTimeout", _
                      cc_isr_Core.UserDefinedErrors.SocketOptionSettingError.Description & _
                      Winsock.LastErrorMessage(" Last error: ")
        Else
            m_receiveTimeout = a_timeoutMs
        End If
    End If
    
End Sub

''' <summary>   Gets the socket receive timeout in milliseconds. </summary>
Public Property Get receiveTimeout() As Long
    
    receiveTimeout = m_receiveTimeout

End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Connection Changed

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal a_isConnected As Boolean)
    
    RaiseEvent ConnectionChanged( _
            cc_isr_Winsock.Constructor.CreateConnectionChangedEventArgs(a_isConnected))

End Sub



