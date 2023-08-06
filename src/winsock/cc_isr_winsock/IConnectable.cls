VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConnectable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''' <summary>   Close the connection and releases the socket. </summary>
Public Sub Dispose()
End Sub

''' <summary>   Opens a socket connection and sets and returns true if connected. </summary>
''' <param name="a_host">        [String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">        [Long] The port that the server is listening on. </param>
''' <param name="a_timeoutMs">   [Optional, Long, 500] The receive timeout in milliseconds. </param>
''' <returns>   [Boolean] True if the socket is connected. </returns>
Public Function OpenConnection(ByVal a_host As String, ByVal a_port As Long, _
        Optional ByVal a_timeoutMs As Long = 500) As Boolean
End Function

''' <summary>   Closes the connection and returns true if not connected. </summary>
''' <returns>   [Boolean] True if disconnection succeeded; otherwise, false. </returns>
Public Function CloseConnection() As Boolean
End Function

''' <summary>   Returns the connection state of the <see cref="IConnectable"/> object. </summary>
''' <value>   [Boolean] True if the socket is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean
End Property

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_eventArgs">   [<see cref="ConnectionChangedEventArgs"/>]. </param>
Public Sub OnConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
End Sub

''' <summary>   Raises the <see cref="ConnectionChanging"/> event. </summary>
''' <para name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>]. </param>
Public Sub OnConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub



