VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IConnectable

Private m_tcpClientSocket As IPv4StreamSocket

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/>] event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   The event that is raised before changing the connection. </summary>
''' <param name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>] event args. </param>
Public Event ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

''' <summary>   Handles the Class Initialize event; constructs the <see cref="TcpClientSocket"/>. </summary>
Private Sub Class_Initialize()
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="TcpClientSocket"/>. </summary>
Private Sub Class_Terminate()
    IConnectable_Dispose
End Sub

''' <summary>   Gets the Tcp Client socket. </summary>
''' <value>   [<see cref="IPv4StreamSocket"/>]. </value>
Public Property Get TcpClientSocket() As IPv4StreamSocket
    Set TcpClientSocket = m_tcpClientSocket
End Property

''' <summary>   Initializes this TCP Client using the provided connected socket. </summary>
''' <remarks>   Raises an invalid operation exception is a socket is already connected. </remarks>
''' <param name="connectedSocket ">   [IPv4StreamSocket] A connected INET stream socket. </param>
Public Sub Initialize(ByVal a_connectedSocket As IPv4StreamSocket)
    
    If Me.Connected Then
        Err.Raise cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code, "cc_isr_Winsock.TcpClient.WinsockApi.Initialize", _
                  cc_isr_Core.UserDefinedErrors.InvalidOperationError.Description & _
                  " This Tc Client is connected to " & Me.TcpClientSocket.SocketAddress & _
                  " A new Tcp Client must be used when initialized using a connected socket at " & _
                  a_connectedSocket.SocketAddress & "."
    End If
    
    ' remove this socket
    If Not m_tcpClientSocket Is Nothing Then Set m_tcpClientSocket = Nothing
    
    Set m_tcpClientSocket = a_connectedSocket
End Sub

''' <summary>   Sends a terminated message to the server. </summary>
''' <param name="a_message">   [String] The message. </param>
''' <returns>   [Long] The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal a_message As String) As Long
    SendMessage = Me.TcpClientSocket.SendMessage(a_message)
End Function

''' <summary>   Receives a message from the server. </summary>
''' <param name="a_maxLength">   [Long] The maximum number of characters to read. </param>
''' <returns>   [String] The received characters. </returns>
Public Function ReceiveBytes(ByVal a_maxLength As Long) As String
    
    ReceiveBytes = Me.TcpClientSocket.ReceiveBytes(a_maxLength)

End Function

''' <summary>   Receives a message from the server. </summary>
''' <remarks>   Data is read until exhausted. </remarks>
''' <param name="a_bufferSize">   [Optional, Integer, 1024] The number of character to read on each trial. </param>
''' <returns>   The received characters. </returns>
Public Function ReceiveString(Optional ByVal a_bufferSize As Integer = 1024) As String
    
    ReceiveString = Me.TcpClientSocket.ReceiveString(a_bufferSize)

End Function

''' <summary>   Receives a message from the server until reaching the specified termination
''' or reading the specified number of characters. </summary>
''' <param name="a_maxLength">     [Optional, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">       [Optional, True] True to return the string without the termination. </param>
''' <param name="a_termination">   [Optional, 10 ASCII] The character the signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function ReceiveRaw(Optional ByVal a_maxLength As Long = &H7FFF, _
                           Optional ByVal a_trimEnd As Boolean = True, _
                           Optional ByVal a_termination As String = vbLf) As String
    
    ReceiveRaw = Me.TcpClientSocket.ReceiveRaw(a_maxLength, a_trimEnd, a_termination)
  
End Function


''' <summary>   Sends a query message and reads back from the server until reaching the specified
''' termination or reading the specified number of characters. </summary>
''' <param name="a_message">       [String] The message. </param>
''' <param name="a_maxLength">     [Optional, Long, 32767] The maximum number of bytes to read. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without the termination. </param>
''' <param name="a_termination">   [Optional, String, 10 ASCII] The character the signifies the end of the stream. </param>
''' <returns>   The received string. </returns>
Public Function QueryRaw(ByVal a_message As String, _
                         Optional ByVal a_maxLength As Long = &H7FFF, _
                         Optional ByVal a_trimEnd As Boolean = True, _
                         Optional ByVal a_termination As String = vbLf) As String
    
    If Me.SendMessage(a_message) > 0 Then
        QueryRaw = Me.TcpClientSocket.ReceiveRaw(a_maxLength, a_trimEnd, a_termination)
    Else
        QueryRaw = vbNullString
    End If
  
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
' Connectable implementation
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean
    CanConnect = IConnectable_CanConnect
End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Private Property Get IConnectable_CanConnect() As Boolean
    If m_tcpClientSocket Is Nothing Then
        IConnectable_CanConnect = False
    Else
        IConnectable_CanConnect = Not m_tcpClientSocket.Connected
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
    
    If m_tcpClientSocket Is Nothing Then
        IConnectable_Connected = False
    Else
        IConnectable_Connected = Me.TcpClientSocket.Connected
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
    
    If p_disconnected And Not m_tcpClientSocket Is Nothing Then Set m_tcpClientSocket = Nothing
    
    On Error GoTo 0

End Sub

''' <summary>   Opens a socket connection and sets and returns true if connected. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_timeoutMs">   [Optional, Long, 500] The receive timeout in milliseconds. </param>
''' <returns>   [Boolean] True if the socket is connected. </returns>
Public Function OpenConnection(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_timeoutMs As Long = 500&) As Boolean
    OpenConnection = IConnectable_OpenConnection(a_host, a_port)
End Function

''' <summary>   Opens a socket connection and sets and returns true if connected. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_timeoutMs">   [Optional, Long, 500] The receive timeout in milliseconds. </param>
''' <returns>   [Boolean] True if the socket is connected. </returns>
Private Function IConnectable_OpenConnection(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_timeoutMs As Long = 500&) As Boolean

    If a_host = "localhost" Then a_host = "127.0.0.1"
    
    Dim p_eventArgs As New cc_isr_Winsock.ConnectionChangingEventArgs
    p_eventArgs.Initialize Me.Connected
    IConnectable_OnConnectionChanging p_eventArgs
    
    If Not p_eventArgs.Cancel Then
    
        ' a new socket must be created on each connection because closing a socket
        ' releases the socket, which can then be used by the system as soon as it is released.
        Set m_tcpClientSocket = New IPv4StreamSocket
    
        Me.TcpClientSocket.OpenConnection a_host, a_port
        
        If Me.TcpClientSocket.Connected Then
            Me.TcpClientSocket.SetReceiveTimeout a_timeoutMs
        End If
        
        OnConnectionChanged Me.Connected
    
    End If
    
    IConnectable_OpenConnection = Me.Connected
    
End Function



''' <summary>   Closes the connection and returns true if not connected. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean
    CloseConnection = IConnectable_CloseConnection()
End Function

''' <summary>   Closes the connection and returns true if not connected. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Private Function IConnectable_CloseConnection() As Boolean

    Dim p_eventArgs As New cc_isr_Winsock.ConnectionChangingEventArgs
    p_eventArgs.Initialize Me.Connected
    IConnectable_OnConnectionChanging p_eventArgs
    
    If Not p_eventArgs.Cancel Then
    
        Me.TcpClientSocket.CloseConnection
        
        ' the socket must be disposed here becuase it cannot be used after
        ' the connection is closed.
        Set m_tcpClientSocket = Nothing
        
        OnConnectionChanged Me.Connected
            
    End If
    
    IConnectable_CloseConnection = Not Me.Connected
    
End Function


''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal a_isConnected As Boolean)
    
    IConnectable_OnConnectionChanged cc_isr_Winsock.Constructor.CreateConnectionChangedEventArgs(a_isConnected)

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
