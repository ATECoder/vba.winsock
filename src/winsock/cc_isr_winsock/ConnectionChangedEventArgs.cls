VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionChangedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type ThisData
    Connected As Boolean
End Type

Private This As ThisData

''' <summary>   Initializes and return the event arguments. </summary>
''' <para name="a_isConnected">   [Boolean, False] The connection state. </param>
''' <returns>   [<see cref="ConnectionChangedEventArgs"/>]. </returns>
Public Function Initialize(Optional ByVal a_isConnected As Boolean = False) As ConnectionChangedEventArgs
    This.Connected = a_isConnected
    Set Initialize = Me
End Function

''' <summary>   Gets the connection state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Connected() As Boolean
    Connected = This.Connected
End Property



