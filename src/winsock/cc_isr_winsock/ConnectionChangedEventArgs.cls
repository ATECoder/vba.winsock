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
Private m_connected As Boolean

''' <summary>   Initializes the event arguments. </summary>
''' <para name="a_isConnected">   [Boolean] The connection state. </param>
Public Sub Initialize(ByVal a_isConnected As Boolean)
    m_connected = a_isConnected
End Sub

''' <summary>   Gets the connection state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Connected() As Boolean
    Connected = m_connected
End Property



