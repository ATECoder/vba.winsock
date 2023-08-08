VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Winsock object  factory. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Instantiates a new <see cref="IPv4StreamSocket"/> object. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="IPv4StreamSocket"/> class. </returns>
Public Function NewIPv4StreamSocket() As IPv4StreamSocket
    Set NewIPv4StreamSocket = New IPv4StreamSocket
End Function

''' <summary>   Instantiates a new <see cref="TcpClient"/> object. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="TcpClient"/>]. </returns>
Public Function NewTcpClient() As TcpClient
    Set NewTcpClient = New TcpClient
End Function

''' <summary>   Instantiates a new <see cref="TcpServer"/> object. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="TcpServer"/]. </returns>
Public Function NewTcpServer() As TcpServer
    Set NewTcpServer = New TcpServer
End Function

''' <summary>   Instantiates a new <see cref="Winsock"/> object. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="Winsock"/]. </returns>
Public Function NewWinsock() As Winsock
    Set NewWinsock = New Winsock
End Function

''' <summary>   Gets an instance of the <see cref="ConnectionChangedEventArgs"/> object. </summary>
''' <value>   [<see cref="ConnectionChangedEventArgs"/>]. </returns>
Public Property Get NewConnectionChangedEventArgs() As ConnectionChangedEventArgs
    Set NewConnectionChangedEventArgs = New ConnectionChangedEventArgs
End Property

''' <summary>   Gets an instance of the <see cref="ConnectionChangingEventArgs"/> object. </summary>
''' <value>   [<see cref="ConnectionChangingEventArgs"/>]. </returns>
Public Property Get NewConnectionChangingEventArgs() As ConnectionChangingEventArgs
    Set NewConnectionChangingEventArgs = New ConnectionChangingEventArgs
End Property



