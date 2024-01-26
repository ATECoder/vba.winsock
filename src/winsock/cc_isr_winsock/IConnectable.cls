VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConnectable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object that can connect. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Gets the address consisting of the IPv4 dotted-decimal host
''' and the port the server is listening too. </summary>
''' <value>   [String] The IPv4 dotted-decimal host address and port delimited with a colon. </value>
Public Property Get Address() As String
End Property

''' <summary>   Returns true if a connection can be made. </summary>
''' <value>   [Boolean] True if a connection can be made. </value>
Public Property Get CanConnect() As Boolean
End Property

''' <summary>   Returns the connection state of the <see cref="IConnectable"/> object. </summary>
''' <value>   [Boolean] True if the object is connected; otherwise, False. </value>
Public Property Get Connected() As Boolean
End Property

''' <summary>   Close the connection and releases this object rsources. </summary>
Public Sub Dispose()
End Sub

''' <summary>   Opens a connection. </summary>
''' <param name="a_address">   [String] An IPv4 dotted-decimal <c>host:port</c> address. </param>
''' <param name="a_timeout">   [Long] The receive timeout in milliseconds. </param>
''' <returns>   [Boolean] True if the connection change successed or cancelled. </returns>
Public Function OpenConnection(ByVal a_address As String, ByVal a_timeout As Long) As Boolean
End Function

''' <summary>   Tries to open a connection. </summary>
''' <param name="a_address">   [String] An IPv4 dotted-decimal <c>host:port</c> address. </param>
''' <param name="a_timeout">   [Long] The receive timeout in milliseconds. </param>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if the connection change successed or cancelled. </returns>
Public Function TryOpenConnection(ByVal a_address As String, _
    ByVal a_timeout As Long, ByRef a_details As String) As Boolean
End Function

''' <summary>   Closes the connection and releases the connection resources. </summary>
''' <returns>   [Boolean] True if the connection change successed or cancelled. </returns>
Public Function CloseConnection() As Boolean
End Function

''' <summary>   Tries to close the connection and release the connection resources. </summary>
''' <param name="a_details">   [Out, String] details the failure. </param>
''' <returns>   [Boolean] True if the connection change successed or cancelled. </returns>
Public Function TryCloseConnection(ByRef a_details As String) As Boolean
End Function

''' <summary>   Raises the <see cref="ConnectionChanged"/> event. </summary>
''' <para name="a_eventArgs">   [<see cref="ConnectionChangedEventArgs"/>]. </param>
Public Sub OnConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)
End Sub

''' <summary>   Raises the <see cref="ConnectionChanging"/> event. </summary>
''' <para name="a_eventArgs">   [<see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/>]. </param>
Public Sub OnConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)
End Sub



