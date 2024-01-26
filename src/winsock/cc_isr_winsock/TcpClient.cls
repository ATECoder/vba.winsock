VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   TCP Client. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Implements IConnectable

Private Type this_
    Socket As IPv4StreamSocket
    Connectable As IConnectable
    Connected As Boolean
End Type

Private This As this_

Private WithEvents m_Socket As IPv4StreamSocket
Attribute m_Socket.VB_VarHelpID = -1

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>] event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before changing the connection. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>] event args. </param>
Public Event ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Handles the Class Initialize event. </summary>
Private Sub Class_Initialize()
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="Socket"/>. </summary>
Private Sub Class_Terminate()
    IConnectable_Dispose
End Sub

''' <summary>   Returns this unconnected Tcp Client instance. No actual initialization is done
''' at this time. </summary>
''' <remarks>   Raises an invalid operation exception is a socket is already connected. </remarks>
''' <returns>   [<see cref="TcpClient"/>]. </returns>
Public Function Initialize() As TcpClient
    
    Set Initialize = Me
    
End Function

''' <summary>   Accepts and active connection using the provided connected socket and returns
''' this instance of a TCP Client. </summary>
''' <remarks>   Raises an invalid operation exception is a socket is already connected. </remarks>
''' <param name="a_connectedSocket ">   [IPv4StreamSocket] A connected INET stream socket. </param>
''' <returns>   [<see cref="TcpClient"/>]. </returns>
Public Function Accept(ByVal a_connectedSocket As IPv4StreamSocket) As TcpClient
    
    If Me.Connected Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".Accept", _
            " This Tcp Client is already connected to " & This.Connectable.Address & _
            " A new Tcp Client must be used to accept a connected socket at " & _
            a_connectedSocket.Address & "."
    End If
    
    ' remove this socket
    If Not This.Socket Is Nothing Then
        Set This.Connectable = Nothing
        Set m_Socket = Nothing
        Set This.Socket = Nothing
    End If
    
    Set This.Socket = a_connectedSocket
    Set m_Socket = This.Socket
    Set This.Connectable = This.Socket
    
    Set Accept = Me
    
End Function

''' <summary>   Gets the Tcp Client socket. </summary>
''' <value>   [<see cref="IPv4StreamSocket"/>]. </value>
Public Property Get Socket() As IPv4StreamSocket
    Set Socket = This.Socket
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  I/O
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sends a terminated message to the server. </summary>
''' <param name="a_message">   [String] The message. </param>
''' <returns>   [Long] The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal a_message As String) As Long
    SendMessage = This.Socket.SendMessage(a_message)
End Function

''' <summary>   Receives a message from the server. </summary>
''' <param name="a_maxLength">     [Optiona, Long, 32767] The maximum number of characters to read. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without the termination. </param>
''' <param name="a_termination">   [Optional, 10 ASCII] The character that signifies the end of the stream. </param>
''' <returns>   [String]. </returns>
Public Function ReceiveBytes(Optional ByVal a_maxLength As Long = 32767, _
    Optional ByVal a_trimEnd As Boolean = True, _
    Optional ByVal a_termination As String = VBA.vbLf) As String
    
    ReceiveBytes = This.Socket.ReceiveBytes(a_maxLength)

End Function

''' <summary>   Receives a message from the server until timeout. </summary>
''' <remarks>   Data is read until exhausted. </remarks>
''' <param name="a_bufferSize">   [Optional, Integer, 1024] The number of character to read on each trial. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveUntilTimeout(Optional ByVal a_bufferSize As Integer = 1024) As String
    
    ReceiveUntilTimeout = This.Socket.ReceiveUntilTimeout(a_bufferSize)

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="a_maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without the termination. </param>
''' <param name="a_termination">   [Optional, 10 ASCII] The character the signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function ReceiveRaw(Optional ByVal a_maxLength As Long = &H7FFF, _
                           Optional ByVal a_trimEnd As Boolean = True, _
                           Optional ByVal a_termination As String = vbLf) As String
    
    ReceiveRaw = This.Socket.ReceiveRaw(a_maxLength, a_trimEnd, a_termination)
  
End Function


''' <summary>   Sends a query message and reads back from the server until reaching the specified
''' termination or reading the specified number of characters. </summary>
''' <param name="a_message">       [String] A terminated query message. </param>
''' <param name="a_maxLength">     [Optional, Long, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without the termination. </param>
''' <param name="a_termination">   [Optional, String, 10 ASCII] The character the signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function QueryRaw(ByVal a_message As String, _
                         Optional ByVal a_maxLength As Long = &H7FFF, _
                         Optional ByVal a_trimEnd As Boolean = True, _
                         Optional ByVal a_termination As String = vbLf) As String
    
    If Me.SendMessage(a_message) > 0 Then
        QueryRaw = This.Socket.ReceiveRaw(a_maxLength, a_trimEnd, a_termination)
    Else
        QueryRaw = vbNullString
    End If
  
End Function

''' <summary>   Receives a message from the server until reaching the specified termination. </summary>
''' <remarks>   The <see cref="ReceiveTimeout"/> must be set long enough the exceed any instrument
''' or controller delays. </remarks>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] detasils the failure reason. </param>
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
    TryReceive = This.Socket.TryReceive(a_buffer, a_details, a_trimEnd, a_termination)
                        
End Function

''' <summary>   Sends a query message and reads back from the server until reaching the specified
''' termination. </summary>
''' <param name="a_message">       [String] A terminated query message. </param>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] detasils the failure reason. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [Integer] If no error, the number of bytes that were received.
''' Otherwise, RECEIVE_ERROR (-1).
''' </returns>
Public Function TryQuery(ByVal a_message As String, _
                         ByRef a_buffer As String, _
                         ByRef a_details As String, _
                         Optional ByVal a_trimEnd As Boolean = True, _
                         Optional ByVal a_termination As String = vbLf) As Long
    If Me.SendMessage(a_message) > 0 Then
        TryQuery = This.Socket.TryReceive(a_buffer, a_details, a_trimEnd, a_termination)
    Else
        TryQuery = vbNullString
    End If
  
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Connectable implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets a reference to the object <see cref="IConnectable"/>. </summary>
''' <value>   [IConnectable]. </value>
Public Property Get Connectable() As IConnectable
    Set Connectable = This.Connectable
End Property

''' <summary>   Gets the address consisting of the IPv4 dotted-decimal host
''' and the port the server is listening too. </summary>
''' <value>   [String] The IPv4 dotted-decimal host address and port delimited with a colon. </value>
Public Property Get Address() As String
    Address = IConnectable_Address
End Property

Private Property Get IConnectable_Address() As String
    If This.Connectable Is Nothing Then
        IConnectable_Address = VBA.vbNullString
    Else
        IConnectable_Address = This.Connectable.Address
    End If
End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean
    CanConnect = IConnectable_CanConnect
End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Private Property Get IConnectable_CanConnect() As Boolean
    If This.Connectable Is Nothing Then
        IConnectable_CanConnect = False
    Else
        IConnectable_CanConnect = This.Connectable.CanConnect
    End If
End Property

''' <summary>   Returns the connection state of the client. </summary>
''' <value>   [Boolean] True if the socket is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
    Connected = IConnectable_Connected
End Property

''' <summary>   Returns the connection state of the client. </summary>
''' <value>   [Boolean] True if the socket is connected; otherwise, False. </value>
Private Property Get IConnectable_Connected() As Boolean
    
    If This.Socket Is Nothing Then
        IConnectable_Connected = False
    Else
        IConnectable_Connected = This.Connectable.Connected And This.Connected
    End If

End Property

''' <summary>   Close the connection and releases the socket. </summary>
Public Sub Dispose()
    IConnectable_Dispose
End Sub

''' <summary>   Close the connection and releases the socket. </summary>
Private Sub IConnectable_Dispose()
    
    On Error Resume Next
    
    Dim p_disconnected As Boolean: p_disconnected = Not Me.Connected
    
    If Not p_disconnected Then p_disconnected = IConnectable_CloseConnection()
    
    Set This.Socket = Nothing
    Set m_Socket = Nothing
    Set This.Connectable = Nothing
    
    On Error GoTo 0

End Sub

''' <summary>   Opens a TCP Client socket connection. </summary>
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
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".IConnectable_OpenConnection", _
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

    Dim p_eventArgs As New cc_isr_Winsock.ConnectionChangingEventArgs
    p_eventArgs.Initialize Me.Connected
    IConnectable_OnConnectionChanging p_eventArgs
    
    If p_eventArgs.Cancel Then
    
        IConnectable_TryOpenConnection = True
        
    Else
    
        ' a new socket must be created on each connection because closing a socket
        ' releases the socket, which can then be used by the system as soon as it is released.
        Set This.Socket = New IPv4StreamSocket
        Set m_Socket = This.Socket
        Set This.Connectable = This.Socket
    
        If This.Connectable.TryOpenConnection(a_address, a_timeout, a_details) Then
    
            IConnectable_TryOpenConnection = True
        
        Else
        
            IConnectable_TryOpenConnection = False
        
        End If
        
    
    End If


End Function

''' <summary>   Closes the connection and returns true if not connected. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean
    CloseConnection = IConnectable_CloseConnection()
End Function

Private Function IConnectable_CloseConnection() As Boolean

    Dim p_details As String: p_details = VBA.vbNullString
    
    If IConnectable_TryCloseConnection(p_details) Then
    
        IConnectable_CloseConnection = True
        
    Else
    
        IConnectable_CloseConnection = False
    
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.SocketDisconnectionError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".IConnectable_CloseConnection", _
            " " & p_details
    
    End If

End Function

''' <summary>   Tries to close the socket connection and zeros the socket id. </summary>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if socket disconnection succeeded; otherwise, false. </returns>
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
    
        If This.Connectable.TryCloseConnection(a_details) Then
        
            IConnectable_TryCloseConnection = True
        
            ' the socket must be disposed here because it cannot be used after
            ' the connection is closed.
            Set This.Socket = Nothing
            m_Socket = Nothing
            Set This.Connectable = Nothing
            
        Else
        
            IConnectable_TryCloseConnection = False
            
        End If
        
    End If

End Function

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal a_isConnected As Boolean)
    
    IConnectable_OnConnectionChanged cc_isr_Winsock.Factory.NewConnectionChangedEventArgs.Initialize(a_isConnected)

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

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  TCP Socket Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Winsock.IPv4SreamSocket.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_Socket_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Const p_procedureName As String = "m_Socket_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    This.Connected = a_eventArgs.Connected
    
    OnConnectionChanged Me.Connected
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' default to false on connection
    This.Connected = False
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the <see cref="cc_isr_Winsock.IPv4SreamSocket.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_Socket_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

    Const p_procedureName As String = "m_Socket_ConnectionChanging"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    IConnectable_OnConnectionChanging a_eventArgs

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' default to false on connection
    This.Connected = False
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub



