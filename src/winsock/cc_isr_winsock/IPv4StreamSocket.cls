VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPv4StreamSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   IPV4 Stream Socket. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const SEND_MESSAGE_ERROR = -1
Private Const SEND_PARTIAL_MESSAGE_ERROR = -2
Private Const RECEIVE_ERROR = -1
Private Const DEFAULT_TIMEOUT_MS = 500

Implements IConnectable

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>] event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before changing the connection. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>] event args. </param>
Public Event ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Permits an incoming connection attempt on a socket. </summary>
''' <param name="a_socketId">             [in] A descriptor that identifies a socket that has been placed in a listening state with the listen function.
'''                                       The connection is actually made with the socket that is returned by accept. </param>
''' <param name="a_clientAddress">        [out] An optional pointer to a buffer that receives the address of the connecting entity,
'''                                        as known to the communications layer. The exact format of the address parameter is determined by the address family
'''                                        that was established when the socket from the ws32_Address structure was created. </param>
''' <param name="a_clientAddressLength">   [in, out] An optional pointer to an integer that contains the length of structure pointed to by
'''                                        the address parameter. </param>
''' <returns>   If no error occurs, accept returns a value of type SOCKET that is a descriptor for the new socket.
''' This returned value is a handle for the socket on which the actual connection is made. Otherwise, a value of
''' ws32_INVALID_SOCKET is returned, and a specific error code can be retrieved by calling WSAGetLastError.
''' The integer referred to by clientAddressLength initially contains the amount of space pointed to by clientAddress. On return it
''' will contain the actual length in bytes of the address returned. <returns>
Private Declare PtrSafe Function accept_ Lib "wsock32.dll" Alias "accept" ( _
        ByVal a_socketId As Long, ByRef a_clientAddress As wsock32.ws32_Address, _
        ByVal a_clientAddressLength As Integer) As Long

''' <summary>   Associates a local address with a socket. </summary>
''' <param name="a_socketid">        [in] A descriptor identifying an unbound socket. </param>
''' <param name="a_address">         [in] A pointer to a ws32_Address_in structure of the local address
'''                                  to assign to the bound socket . </param>
''' <param name="a_addressLength">   [in] The length, in bytes, of the value pointed to by address. </param>
''' <returns>   Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function bind_ Lib "wsock32.dll" Alias "bind" ( _
        ByVal a_socketId As Long, ByRef a_address As wsock32.ws32_Address_in, _
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
''' <param name="a_address">      [in] A pointer to the <see cref="ws32_Address_in"/>   structure to which the
'''                               connection should be established. </param>
''' <param name="a_addressLen">   [in] The length, in bytes, of the ws32_Address structure pointed to by the
'''                               <paramref name="address"/>   parameter. </param>
''' <returns>   Zero if no error occurs; otherwise, a value of ws32_SOCKET_ERROR is returned.
''' A specific error code can be retrieved by calling WSAGetLastError. <returns>
Private Declare PtrSafe Function connect_ Lib "wsock32.dll" Alias "connect" ( _
        ByVal a_socketId As Long, ByRef a_address As wsock32.ws32_Address_in, _
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
''' <param name="a_optionValue">    [in] A pointer to the buffer in which the value for the requested option is specified. </param>
''' <param name="a_optionLength">   [in] The size, in bytes, of the buffer pointed to by the
'''                                 <paramref="a_optionValue"/> argument. </param>
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

Private Type this_
    Port As Long
    Host As String
    Address As String
    SocketId As Long
    Connected As Boolean
    Listening As Boolean
    ReceiveTimeout As Long
    FdSet As wsock32.ws32_fd_set
End Type

Private This As this_

''' <summary>   Initializes an internet protocol Stream Socket with a socket id. </summary>
Private Sub Class_Initialize()

    ' see that the singleton Winsock class is instantiated
    On Error Resume Next
    
    ' this is required to initialize Winsock.  It will only ran once.
    Winsock.Initialize
    
    This.Host = vbNullString
    This.Address = vbNullString
    
    On Error GoTo 0

    ' Create a new socket
    
    This.SocketId = WinsockApi.CreateIPv4StreamSocket()
    If This.SocketId = wsock32.ws32_INVALID_SOCKET Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketCreationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".Initialize", _
            " Failed creating socket. winsock32.CreateIPv4StreamSocket() returned invalid socket ID of " & _
            VBA.CStr(This.SocketId) & Winsock.LastErrorMessage()
    Else
        Winsock.Register Me
        
        ' set the receive timeout to its default value
        
        Me.ReceiveTimeoutSetter DEFAULT_TIMEOUT_MS
    End If
    
End Sub

''' <summary>   Terminates the socket and releases its Winsock registration. </summary>
Private Sub Class_Terminate()

    ' see that the singleton Winsock class is instantiated
    On Error Resume Next
    
    ' unregister this socket.
    Winsock.Unregister Me
    
    On Error GoTo 0

End Sub

''' <summary>   Returns this unconnected Inet stream socket instance. No actual initialization is done
''' at this time. </summary>
''' <returns>   [<see cref="IPv4StreamSocket"/>]. </returns>
Public Function Initialize() As IPv4StreamSocket
    
    Set Initialize = Me

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Socket implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns the port the server is listening to. </summary>
''' <value>   [Long] The port the server listens to. </value>
Public Property Get Port() As Long
    Port = This.Port
End Property

''' <summary>   Gets the IPv4 dotted-decimal host address. </summary>
''' <value>   [String] The IPv4 dotted-decimal host address. </value>
Public Property Get Host() As String
    Host = This.Host
End Property

''' <summary>   Returns a descriptor identifying the socket. </summary>
''' <value>   The socket identifying descriptor. </value>
Public Property Get SocketId() As Long
Attribute SocketId.VB_Description = "A descriptor identifying the socket"
Attribute SocketId.VB_UserMemId = 0
    SocketId = This.SocketId
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Connectable implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the address consisting of the IPv4 dotted-decimal host
''' and the port the server is listening too. </summary>
''' <value>   [String] The IPv4 dotted-decimal host address and port delimited with a colon. </value>
Public Property Get Address() As String
    Address = IConnectable_Address
End Property

Private Property Get IConnectable_Address() As String
    IConnectable_Address = This.Address
End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean
    CanConnect = IConnectable_CanConnect
End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Private Property Get IConnectable_CanConnect() As Boolean
    
    IConnectable_CanConnect = This.SocketId <> wsock32.ws32_INVALID_SOCKET
    
End Property

''' <summary>   Returns the connection state of the socket. </summary>
''' <value>   True if the socket is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
    
    Connected = IConnectable_Connected

End Property

Private Property Get IConnectable_Connected() As Boolean
    
    IConnectable_Connected = This.Connected

End Property

''' <summary>   Close the connection and releases the socket resources. </summary>
Public Sub Dispose()
    IConnectable_Dispose
End Sub

Private Sub IConnectable_Dispose()
    
    On Error Resume Next
    
    Dim p_disconnected As Boolean: p_disconnected = Not Me.Connected
    
    If Not p_disconnected Then p_disconnected = IConnectable_CloseConnection()
    
    This.SocketId = wsock32.ws32_INVALID_SOCKET
    
    On Error GoTo 0

End Sub

''' <summary>   Opens a socket connection. </summary>
''' <param name="a_address">   [String] An IPv4 dotted-decimal <c>host:port</c> address. </param>
''' <returns>   [Boolean] True if the connection change successed or cancelled. </returns>
Public Function OpenConnection(ByVal a_address As String, ByVal a_timeout As Long) As Boolean
    OpenConnection = IConnectable_OpenConnection(a_address, a_timeout)
End Function

Private Function IConnectable_OpenConnection(ByVal a_address As String, ByVal a_timeout As Long) As Boolean
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    If IConnectable_TryOpenConnection(a_address, a_timeout, p_details) Then
    
        IConnectable_OpenConnection = True
        
    Else
    
        IConnectable_OpenConnection = False
    
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".OpenConnection", _
            " " & p_details
    
    End If

End Function

''' <summary>   Tries to open a socket connection. </summary>
''' <param name="a_address">   [String] An IPv4 dotted-decimal <c>host:port</c> address. </param>
''' <param name="a_timeout">   [Long] The receive timeout in milliseconds. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if the connection change successed or cancelled. </returns>
Public Function TryOpenConnection(ByVal a_address As String, _
    ByVal a_timeout As Long, ByRef a_details As String) As Boolean
        
    TryOpenConnection = IConnectable_TryOpenConnection(a_address, a_timeout, a_details)
    
End Function

Private Function IConnectable_TryOpenConnection(ByVal a_address As String, _
    ByVal a_timeout As Long, ByRef a_details As String) As Boolean

    If cc_isr_Winsock.Winsock.TryParseAddress(a_address, This.Host, This.Port, a_details) Then
    
        Dim p_eventArgs As New cc_isr_Winsock.ConnectionChangingEventArgs
        p_eventArgs.Initialize Me.Connected
        IConnectable_OnConnectionChanging p_eventArgs
    
        If p_eventArgs.Cancel Then
        
            IConnectable_TryOpenConnection = True
            
        Else
        
            This.Connected = False
            
            Dim p_Address As wsock32.ws32_Address_in
            
            ' set the binary address
            p_Address.sin_addr.s_addr = WinsockApi.ToInetAddress(This.Host)
            
            p_Address.sin_family = wsock32.ws32_AF_INET
            
            ' set the port in network byte order
            p_Address.sin_port = WinsockApi.ToInetByteOrder(This.Port)
            
            Dim p_connectResult As Long
            p_connectResult = connect_(This.SocketId, p_Address, Len(p_Address))
            
            If p_connectResult = wsock32.ws32_SOCKET_ERROR Then
                IConnectable_TryOpenConnection = False
                This.Connected = False
                a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".IConnectable_TryOpenConnection" & _
                    ". Failed opening a connection to " & This.Host & ":" & VBA.CStr(This.Port) & _
                    "; winsock32.connect(" & VBA.CStr(This.SocketId) & ", address, length) " & _
                    "returned socket error " & _
                    VBA.CStr(p_connectResult) & Winsock.LastErrorMessage()
            Else
                This.Address = a_address
                This.Connected = True
                Me.ReceiveTimeoutSetter a_timeout
                IConnectable_TryOpenConnection = True
                OnConnectionChanged This.Connected
            End If
            
        End If
    
    Else
    
    IConnectable_TryOpenConnection = False
    
    End If
    
End Function

''' <summary>   Tries to close the socket connection and zeros the socket id. </summary>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if the connection change successed or cancelled. </returns>
Public Function TryCloseConnection(ByRef a_details As String) As Boolean
    TryCloseConnection = IConnectable_TryCloseConnection(a_details)
End Function

Private Function IConnectable_TryCloseConnection(ByRef a_details As String) As Boolean

    Dim p_eventArgs As New cc_isr_Winsock.ConnectionChangingEventArgs
    p_eventArgs.Initialize Me.Connected
    IConnectable_OnConnectionChanging p_eventArgs
    
    If p_eventArgs.Cancel Then
    
        IConnectable_TryCloseConnection = True
    
    Else
    
        Dim p_result As Long
        p_result = closesocket_(This.SocketId)
        
        ' set the socket id to invalid to indicate that it is no longer viable.
        This.SocketId = wsock32.ws32_INVALID_SOCKET
        
        ' the socket is no longer valid; this is disconnected
        This.Connected = False
        
        If p_result = wsock32.ws32_SOCKET_ERROR Then
            
            IConnectable_TryCloseConnection = False
            
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".IConnectable_TryCloseConnection" & _
                " failed closing a connection to " & This.Address & _
                "; winsock32.closesocket(" & VBA.CStr(This.SocketId) & ") " & _
                "returned socket error " & _
                VBA.CStr(p_result) & Winsock.LastErrorMessage()
                    
        Else
        
            IConnectable_TryCloseConnection = True
            
            OnConnectionChanged This.Connected
            
        End If
            
    End If

End Function

''' <summary>   Closes the socket connection and zeros the socket id. </summary>
''' <returns>   [Boolean] True if the connection change successed or cancelled. </returns>
Public Function CloseConnection() As Boolean
    CloseConnection = IConnectable_CloseConnection
End Function

''' <summary>   Closes the socket connection and zeros the socket id. </summary>
''' <returns>   [Boolean] True if the connection change successed or cancelled. </returns>
Private Function IConnectable_CloseConnection() As Boolean

    Dim p_details As String: p_details = VBA.vbNullString
    
    If IConnectable_TryCloseConnection(p_details) Then
        
        IConnectable_CloseConnection = True
    
    Else
    
        IConnectable_CloseConnection = False
    
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketDisconnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".CloseConnection", _
            p_details
    
    End If
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Send
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sends a terminated message to the server. </summary>
''' <param name="a_message">   [String] The terminated message. </param>
''' <param name="a_delay">     [Optional, Long, 1] read after write delay in milliseconds. </param>
''' <returns>   [Long} The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal a_message As String, Optional ByVal a_delay As Long = 1) As Long

    If VBA.vbNullString = a_message Then
        SendMessage = 0
    Else
        Dim p_messageLength As Integer: p_messageLength = Len(a_message)
        Dim p_sentLength As Integer
        p_sentLength = send_(This.SocketId, ByVal a_message, Len(a_message), 0)
        
        If p_sentLength < 0 Then
            
            ' if the sent length is negative, we have an API error.
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".SendMessage", _
                " Failed sending " & a_message & _
                "; winsock32.send(" & VBA.CStr(This.SocketId) & ", message, length, 0) " & _
                "returned socket error " & _
                VBA.CStr(p_sentLength) & Winsock.LastErrorMessage()
            
            SendMessage = SEND_MESSAGE_ERROR
            
        ElseIf p_sentLength <> p_messageLength Then
            
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketSendError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".SendMessage", _
                " Failed sending " & a_message & _
                "; winsock32.send(" & VBA.CStr(This.SocketId) & ", message, length, 0) " & _
                "; A partial message was sent: " & VBA.CStr(p_sentLength) & " out of " & VBA.CStr(p_messageLength)
            SendMessage = SEND_PARTIAL_MESSAGE_ERROR
            
        Else
            SendMessage = p_sentLength
        End If
        
        cc_isr_Core_IO.Factory.NewStopwatch().Wait a_delay
        
    End If
    

End Function

''' <summary>   Sends a terminated message to the server. </summary>
''' <param name="a_message">   [String] The terminated message. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <param name="a_delay">     [Optional, Long, 1] read after write delay in milliseconds. </param>
''' <returns>   [Long]. 0 if nothing to send, negative (-1) if error,
''' positive if the entire message was sent. </returns>
Public Function TrySendMessage(ByVal a_message As String, _
    ByRef a_details As String, _
    Optional ByVal a_delay As Long = 1) As Long

    If VBA.vbNullString = a_message Then
    
        TrySendMessage = 0
        
    Else
    
        Dim p_messageLength As Integer: p_messageLength = Len(a_message)
        Dim p_sentLength As Integer
        p_sentLength = send_(This.SocketId, ByVal a_message, Len(a_message), 0)
        
        If p_sentLength < 0 Then
            
            ' if the sent length is negative, we have an API error.
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".SendMessage" & _
                ". Failed sending '" & a_message & "' to '" & This.Address & _
                "'; winsock32.send(" & VBA.CStr(This.SocketId) & ", message, length, 0) " & _
                "returned socket error " & _
                VBA.CStr(p_sentLength) & Winsock.LastErrorMessage()
            
            TrySendMessage = p_sentLength
            
        ElseIf p_sentLength <> p_messageLength Then
            
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".SendMessage" & _
                ". Failed sending '" & a_message & "' to '" & This.Address & _
                "'; winsock32.send(" & VBA.CStr(This.SocketId) & ", message, length, 0) " & _
                "; A partial message was sent: " & VBA.CStr(p_sentLength) & " out of " & VBA.CStr(p_messageLength)
                
            TrySendMessage = -1
            
        Else
            
            TrySendMessage = p_sentLength
        
        End If
        
        cc_isr_Core_IO.Factory.NewStopwatch().Wait a_delay
        
    End If

End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Receive
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Receives a null terminated message from the server until reaching
''' the specified termination or the maximum length. </summary>
''' <param name="a_buffer">        [String] Where the message will go. </param>
''' <param name="a_maxLength">     [Integer] The maximum number of bytes to read. </param>
''' <param name="a_termination">   [Optional, String, vbLf] The character that signifies the end of the stream. </param>
''' <returns>   [Integer] If no error, the number of byte that were received.
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
        p_count = recv_(This.SocketId, c, l, 0)
        
        If p_count < 1 Then
            ReceiveTerminatedMessage = RECEIVE_ERROR
            a_buffer = VBA.Chr$(0)
            Exit Function
        End If
        
        If c = a_termination Then
           a_buffer = a_buffer + VBA.Chr$(0)
           ReceiveTerminatedMessage = p_length
           Exit Function
        End If
        
        p_length = p_length + p_count
        a_buffer = a_buffer + c
        
    Wend
    
    ReceiveTerminatedMessage = RECEIVE_ERROR
    
End Function

''' <summary>   Receives a message from the server until reaching the specified termination. </summary>
''' <remarks>   The <see cref="ReceiveTimeout"/> must be set long enough the exceed any instrument
''' or controller delays. </remarks>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [Integer] If no error, the number of bytes that were received.
''' Otherwise, RECEIVE_ERROR (-1).
''' </returns>
Public Function TryReceive(ByRef a_buffer As String, _
                           ByRef a_details As String, _
                           Optional ByVal a_trimEnd As Boolean = True, _
                           Optional ByVal a_termination As String = vbLf) As Long

    Dim c As String * 1
    Dim p_count As Long
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder.Initialize()
    Dim p_currentLength As Long: p_currentLength = 0
    a_details = VBA.vbNullString
    
    Dim p_done As Boolean
    Do While Not p_done
    
        DoEvents
        
        c = VBA.vbNullString
        p_count = recv_(This.SocketId, c, 1, 0)
        
        If p_count < 1 Then
        
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".TryReceiveRaw" & _
                ". Failed receiving from '" & This.Address & _
                "'; winsock32.recv(" & VBA.CStr(This.SocketId) & ", buffer, 1, 0) returned " & _
                VBA.CStr(p_count) & Winsock.LastErrorMessage()
            p_currentLength = RECEIVE_ERROR
            p_done = True
            
        ElseIf c = a_termination Then
            
            If Not a_trimEnd Then
                p_builder.Append c
                p_currentLength = p_currentLength + p_count
            End If
            p_done = True
        
        Else
            ' ignore a received CR character if trim end
            If Not (a_trimEnd And c = VBA.vbCr) Then
                p_currentLength = p_currentLength + p_count
                p_builder.Append c
            End If
        End If
    Loop
    
    a_buffer = p_builder.ToString
    TryReceive = p_currentLength

End Function

''' <summary>   Sends a terminated query message and reads back from the server until reaching the specified
''' termination. </summary>
''' <param name="a_message">       [String] The unterminated query message. </param>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] detasils the failure reason. </param>
''' <param name="a_delay">         [Optional, Long, 1] read after write delay in milliseconds. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [Integer] If no error, the number of bytes that were received.
''' 0 if nothing was sent; negative if a sent error occurred, such as mismatch between sent
''' count and message length. Otherwise, RECEIVE_ERROR (-1).
''' </returns>
Public Function TryQuery(ByVal a_message As String, _
                         ByRef a_buffer As String, _
                         ByRef a_details As String, _
                         Optional ByVal a_delay As Long = 1, _
                         Optional ByVal a_trimEnd As Boolean = True, _
                         Optional ByVal a_termination As String = VBA.vbLf) As Long
                         
    Dim p_sentLength As Long
    p_sentLength = Me.TrySendMessage(a_message, a_details, a_delay)
    If p_sentLength = 0 Then
        a_buffer = VBA.vbNullString
        a_details = "Query command is empty."
        TryQuery = 0
    ElseIf p_sentLength > 0 Then
        TryQuery = Me.TryReceive(a_buffer, a_details, a_trimEnd, a_termination)
    Else
        TryQuery = p_sentLength
    End If
  
End Function


''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="a_maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without the termination. </param>
''' <param name="a_termination">   [Optional, 10 ASCII] The character that signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function ReceiveRaw(Optional ByVal a_maxLength As Long = &H7FFFF, _
                           Optional ByVal a_trimEnd As Boolean = True, _
                           Optional ByVal a_termination As String = vbLf) As String
    Dim c As String * 1
    Dim p_count As Integer
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder.Initialize()
    Dim p_currentLength As Long: p_currentLength = 0
    
    Do While p_currentLength < a_maxLength
    
        DoEvents
        
        c = vbNullString
        p_count = recv_(This.SocketId, c, 1, 0)
        
        If p_count < 1 Then
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".ReceiveRaw", _
                " winsock32.recv(" & VBA.CStr(This.SocketId) & ", buffer, 1, 0) returned " & _
                VBA.CStr(p_count) & Winsock.LastErrorMessage()
            Exit Do
        ElseIf c = a_termination Then
            If Not a_trimEnd Then
                p_builder.Append c
            End If
            Exit Do
        Else
            ' ignore a received CR character if trim end
            If Not (a_trimEnd And c = VBA.vbCr) Then
                p_currentLength = p_currentLength + p_count
                p_builder.Append c
            End If
        End If
    Loop
    ReceiveRaw = p_builder.ToString
    
End Function

''' <summary>   Receives a message from the server. </summary>
''' <param name="a_maxLength">     [Optiona, Long, 32767] The maximum number of characters to read. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without the termination. </param>
''' <param name="a_termination">   [Optional, 10 ASCII] The character that signifies the end of the stream. </param>
''' <returns>   [String]. </returns>
Public Function ReceiveBytes(Optional ByVal a_maxLength As Long = 32767, _
    Optional ByVal a_trimEnd As Boolean = True, _
    Optional ByVal a_termination As String = VBA.vbLf) As String
    
    Dim p_buffer As String: p_buffer = StringExtensions.Repeat(VBA.Chr$(0), a_maxLength)
    
    Dim p_actualLength As Long
    
    p_actualLength = recv_(This.SocketId, p_buffer, a_maxLength, 0)
    
    If p_actualLength > 0 Then
        p_buffer = StringExtensions.Substring(p_buffer, 0, p_actualLength)
        If a_trimEnd Then
            p_buffer = cc_isr_Core.StringExtensions.TrimEnd(p_buffer, a_termination)
        End If
        ReceiveBytes = p_buffer
    Else
        ReceiveBytes = vbNullString
    End If
    
End Function

''' <summary>   Receives a message from the server until nothign to receive as indicated by a timeout. </summary>
''' <remarks>   Data is read until exhausted. </remarks>
''' <param name="a_bufferSize">   [Optional, 1024] The number of character to read on each trial. </param>
''' <returns>   [String] The received characters. </returns>
Public Function ReceiveUntilTimeout(Optional ByVal a_bufferSize As Integer = 1024) As String
    
    Dim p_buffer As String
    Dim p_builder As cc_isr_Core.StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder.Initialize()
    
    Do
        p_buffer = Trim(Me.ReceiveBytes(a_bufferSize, False))
        If Len(p_buffer) > 0 Then p_builder.Append p_buffer
        
    Loop While Len(p_buffer) > 0
    
    ReceiveUntilTimeout = p_builder.ToString()
    
End Function

''' <summary>   Sets the socket receive timeout. </summary>
''' <param name="a_timeout">   The timeout interval in milliseconds. </param>
Public Sub ReceiveTimeoutSetter(ByVal a_timeout As Long)

    Dim p_result As Long
    
    ' check if the requested receive timeout is different from the
    ' timeout that was already set.
    If This.ReceiveTimeout <> a_timeout Then
    
        p_result = setsockopt_(This.SocketId, wsock32.ws32_SOL_SOCKET, wsock32.ws32_SO_RCVTIMEO, a_timeout, 4)
    
        If p_result = wsock32.ws32_SOCKET_ERROR Then
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketReceiveError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".ReceiveTimeoutSetter", _
                " Failed setting receive timeout. winsock32.setsockopt(" & _
                VBA.CStr(This.SocketId) & ", wsock32.ws32_SOL_SOCKET, wsock32.ws32_SO_RCVTIMEO, " & _
                VBA.CStr(a_timeout) & ", 4) returned " & _
                VBA.CStr(p_result) & Winsock.LastErrorMessage()
        Else
            This.ReceiveTimeout = a_timeout
        End If
    End If
    
End Sub

''' <summary>   Gets the socket receive timeout in milliseconds. </summary>
''' <value>   [Long]. </value>
Public Property Get ReceiveTimeout() As Long
    
    ReceiveTimeout = This.ReceiveTimeout

End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' TCP Client implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Accepts an active connection, sets the provided socket read timeout, and
''' returns this object as an active Inet stream socket using the provided socket id. </summary>
''' <param name="a_connectedSocketId">   [Long] The ID of a connected socket. </param>
''' <param name="a_timeout">           [Optional, Long, 500] The receive timeout interval in milliseconds. </param>
''' <returns>   [<see cref="IPv4StreamSocket"/>]. </returns>
Public Function Accept(ByVal a_connectedSocketId As Long, _
        Optional ByVal a_timeout As Long = 500) As IPv4StreamSocket
    
    If Me.Connected Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".Accept", _
            " This Inet Stream Socket is already connected to " & This.Address & _
            " A new INet Stream Socket must be used to accept this connected socket ID " & _
            a_connectedSocketId & "."
    End If
    
    
    This.SocketId = a_connectedSocketId
    This.Connected = True
    Me.ReceiveTimeoutSetter a_timeout
    Set Accept = Me

End Function

''' <summary>   Associates a local address with a socket. </summary>
''' <param name="a_address">   The INet address to bind to. </param>
''' <param name="a_port">   The port that the server is listening on. </param>
''' <returns>   [Boolean] True if the socket is bound. </returns>
Public Function BindTo(ByVal a_address As Long, ByVal a_port As Long) As Boolean

    Dim p_endpoint As wsock32.ws32_Address_in
    p_endpoint.sin_family = wsock32.ws32_AF_INET
    p_endpoint.sin_addr.s_addr = a_address
    p_endpoint.sin_port = WinsockApi.ToInetByteOrder(a_port)
    
    Dim p_bindResult As Long
    p_bindResult = bind_(This.SocketId, p_endpoint, 16)
    
    If p_bindResult = wsock32.ws32_SOCKET_ERROR Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketDisconnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".BindTo", _
            " Failed binding to address " & VBA.CStr(a_address) & " at port " & VBA.CStr(a_port) & _
            "; winsock32.bind(" & VBA.CStr(This.SocketId) & ", endpoint, 16) " & _
            "returned socket error " & _
            VBA.CStr(p_bindResult) & Winsock.LastErrorMessage()
        This.Connected = False
    Else
        This.Connected = True
    End If
    BindTo = This.Connected
   
End Function

''' <summary>   Returns the listening state of the socket. </summary>
''' <value>   [Boolean] True if the socket is Listening; otherwise, False. </value>
Public Property Get Listening() As Boolean
    Listening = This.Listening
End Property


''' <summary>   Places a socket in a state in which it is listening for an incoming connection. </summary>
''' <param name="a_backlog">   [Optional, 10] [in] The maximum length of the queue of pending connections. </para>
''' <returns>   [Boolean] True if the socket is Listening; otherwise, False. </returns>
Public Function StartListening(Optional ByVal a_backlog As Integer = 10) As Boolean

    Dim p_listenResult As Long
    p_listenResult = listen_(This.SocketId, a_backlog)
    
    If p_listenResult = wsock32.ws32_SOCKET_ERROR Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketListenError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".StartListening", _
            " Failed to start listening at socket " & VBA.CStr(This.SocketId) & _
            "; winsock32.listen(" & VBA.CStr(This.SocketId) & ", " & VBA.CStr(a_backlog) & ") " & _
            "returned socket error " & _
            VBA.CStr(p_listenResult) & Winsock.LastErrorMessage()
        This.Listening = False
    Else
        This.Listening = True
    End If
    StartListening = This.Listening

End Function

''' <summary>   Permits an incoming connection attempt on a socket. </summary>
''' <param name="a_serverSocket">   [<see cref="IPv4StreamSocket"/>] The server socket. </param>
''' <param name="a_timeout">      [Optional, Long, 500] The maximum milliseconds time for select to wait. </param>
''' <returns>   the connected socket. </returns>
Public Function AcceptTcpClient(ByVal a_serverSocket As IPv4StreamSocket, _
        Optional ByVal a_timeout As Long = 500) As IPv4StreamSocket

    ' add this socket to a socket set.
    WinsockApi.FD_SET_INIT This.FdSet
    WinsockApi.FD_SET_ADD This.SocketId, This.FdSet
    
    ' check if we have available socket for readability
    
    Dim p_readableSocketCount As Integer
    p_readableSocketCount = WinsockApi.DetermineReadability(This.FdSet, a_timeout)
    
    If p_readableSocketCount = 0 Then
        Set AcceptTcpClient = Nothing
        Exit Function
    End If
    
    Dim p_socketId As Long
    Dim p_Address As wsock32.ws32_Address
    
    ' get a connected socket id
    
    p_socketId = accept_(a_serverSocket.SocketId, p_Address, wsock32.ws32_AddressLen)
    
    If p_socketId = wsock32.ws32_INVALID_SOCKET Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketListenError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".AcceptTcpClient", _
            " Failed accepting a tcp client at server socket " & VBA.CStr(a_serverSocket.SocketId) & _
            "; winsock32.accept(" & VBA.CStr(a_serverSocket.SocketId) & ", Address, length) " & _
            "returned invalid socket error " & _
            VBA.CStr(p_socketId) & Winsock.LastErrorMessage()
        Exit Function
    End If
    
    Dim p_clientSocket As IPv4StreamSocket
    Set p_clientSocket = New IPv4StreamSocket
    p_clientSocket.Accept p_socketId, a_timeout
    
    Set AcceptTcpClient = p_clientSocket
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Connection Changed
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal a_isConnected As Boolean)
    
    IConnectable_OnConnectionChanged _
        cc_isr_Winsock.Factory.NewConnectionChangedEventArgs.Initialize(a_isConnected)

End Sub

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_eventArgs">   [<<see cref="ConnectionChangedEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanged(ByVal a_eventArgs As ConnectionChangedEventArgs)
    
    RaiseEvent ConnectionChanged(a_eventArgs)

End Sub

''' <summary>   Raises the <see cref="ConnectionChanging"/> event. </summary>
''' <para name="a_eventArgs">   [<<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>]. </param>
Private Sub IConnectable_OnConnectionChanging(ByVal a_eventArgs As ConnectionChangingEventArgs)
    
    RaiseEvent ConnectionChanging(a_eventArgs)

End Sub


