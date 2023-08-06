VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Constructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' <summary>   Creates a new instance of the <see cref="IPv4StreamSocket"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="IPv4StreamSocket"/> class. </returns>
Public Function CreateIPv4StreamSocket() As IPv4StreamSocket
    Set CreateIPv4StreamSocket = New IPv4StreamSocket
End Function

''' <summary>   Creates a new instance of the <see cref="TcpClient"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="TcpClient"/> class. </returns>
Public Function CreateTcpClient() As TcpClient
    Set CreateTcpClient = New TcpClient
End Function

''' <summary>   Creates a new instance of the <see cref="TcpServer"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="TcpServer"/> class. </returns>
Public Function CreateTcpServer() As TcpServer
    Set CreateTcpServer = New TcpServer
End Function

''' <summary>   Creates a new instance of the <see cref="Winsock"/> class to
''' share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="Winsock"/> class. </returns>
Public Function CreateWinsock() As Winsock
    Set CreateWinsock = New Winsock
End Function

''' <summary>   Gets an instance of the <see cref="ConnectionChangedEventArgs"/> class. </summary>
''' <para name="isConnected">   [Optional, Boolean, False] True if connected. </param>
''' <value>   [<see cref="ConnectionChangedEventArgs"/>]. </returns>
Public Property Get CreateConnectionChangedEventArgs(Optional ByVal a_isConnected As Boolean = False) As ConnectionChangedEventArgs
    Set CreateConnectionChangedEventArgs = New ConnectionChangedEventArgs
    CreateConnectionChangedEventArgs.Initialize a_isConnected
End Property

''' <summary>   Gets an instance of the <see cref="ConnectionChangingEventArgs"/> class. </summary>
''' <para name="isConnected">   [Optional, Boolean, False] True if connected. </param>
''' <value>   [<see cref="ConnectionChangingEventArgs"/>]. </returns>
Public Property Get CreateConnectionChangingEventArgs(Optional ByVal a_isConnected As Boolean = False) As ConnectionChangingEventArgs
    Set CreateConnectionChangingEventArgs = New ConnectionChangingEventArgs
    CreateConnectionChangingEventArgs.Initialize a_isConnected
End Property



