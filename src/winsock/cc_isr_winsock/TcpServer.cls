VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   TCP Server. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    TcpServerSocket As IPv4StreamSocket
End Type

Private This As this_

''' <summary>   The event that is raised upon change of connection. </summary>
''' <param name="a_eventArgs">   An <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event args. </param>
Public Event ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

''' <summary>   Handles the Class Initialize event; constructs the <see cref="TcpServerSocket"/>. </summary>
Private Sub Class_Initialize()
    
    Set This.TcpServerSocket = New IPv4StreamSocket

End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="TcpServerSocket"/>. </summary>
Private Sub Class_Terminate()
    
    Me.Dispose

End Sub

''' <summary>   Gets the Tcp Server socket. </summary>
''' <value>   [<see cref="IPv4StreamSocket"/>]. </value>
Public Property Get TcpServerSocket() As IPv4StreamSocket
    
    Set TcpServerSocket = This.TcpServerSocket

End Property

''' <summary>   Associates any local address with a socket and start listening. </summary>
''' <param name="a_port">      [Long] The port that the server is listening on. </param>
''' <param name="a_backlog">   [Optional, Integer, 10] The maximum length of the queue of pending connections. </para>
''' <returns>   [Boolean] True if the socket is bound. </returns>
Public Function BindTo(ByVal a_port As Long, Optional ByVal a_backlog As Integer = 10) As Boolean

    On Error GoTo ErrHandler
    
    ' bind the socket to the specified port for listening on any address
    ' an exception is raised on failure.
    
    Me.TcpServerSocket.BindTo wsock32.ws32_INADDR_ANY, a_port
    
    OnConnectionChanged Me.Connected
    
    ' get the socket to start listening
    ' an exception is raised on failure.
    
    BindTo = Me.TcpServerSocket.StartListening(a_backlog)
    
ExitHandler:

    Exit Function
    
ErrHandler:

    If Me.Connected Then
        Me.TcpServerSocket.CloseConnection
    End If
    
    OnConnectionChanged Me.Connected
    
End Function

''' <summary>   Returns the connection state of the client. </summary>
''' <returns>   [Boolean] True if the socket is connected; otherwise, False. </returns>
Public Property Get Connected() As Boolean
    
    If This.TcpServerSocket Is Nothing Then
        Connected = False
    Else
        Connected = Me.TcpServerSocket.Connected
    End If

End Property

''' <summary>   Permits an incoming connection attempt on a socket. </summary>
''' <param name="a_timeout">   [Optional, 500] The maximum milliseconds time for select to wait. </param>
''' <returns>   the connected Tcp client. </returns>
Public Function AcceptTcpClient(Optional ByVal a_timeout As Long = 500) As TcpClient

    Dim p_connectedClientSocket As IPv4StreamSocket
    Set p_connectedClientSocket = Me.TcpServerSocket.AcceptTcpClient(Me, a_timeout)
    
    Dim p_client As TcpClient
    Set p_client = New TcpClient
    p_client.Accept p_connectedClientSocket
    
    Set AcceptTcpClient = p_client

End Function

''' <summary>   Closes and releases the socket. </summary>
Public Sub Dispose()
    
    If Me.Connected Then
        Me.TcpServerSocket.CloseConnection
    End If
    Set This.TcpServerSocket = Nothing

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
' Connection Changed
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_isConnected">   [Boolean] The connection state. </param>
Private Sub OnConnectionChanged(ByVal a_isConnected As Boolean)
    
    RaiseEvent ConnectionChanged( _
            cc_isr_Winsock.Factory.NewConnectionChangedEventArgs.Initialize(a_isConnected))

End Sub





