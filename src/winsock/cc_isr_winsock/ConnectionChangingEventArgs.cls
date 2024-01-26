VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionChangingEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Connection changing event arguments. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Connected As Boolean
    Cancel As Boolean
End Type

Private This As this_

''' <summary>   Initializes and returns the event arguments. </summary>
''' <para name="isConnected">   [Boolean, False] The connection state. </param>
''' <returns>   [<see cref="ConnectionChangedEventArgs"/>]. </returns>
Public Function Initialize(Optional ByVal a_isConnected As Boolean = False) As ConnectionChangingEventArgs
    This.Connected = a_isConnected
    This.Cancel = False
    Set Initialize = Me
End Function

''' <summary>   Gets the connection state. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Connected() As Boolean
    Connected = This.Connected
End Property

''' <summary>   Gets the cancel request. </summary>
''' <value>   [Boolean]. </value>
Public Property Get Cancel() As Boolean
    Cancel = This.Cancel
End Property

''' <summary>   Sets the cancel request. </summary>
''' <param name="a_value">   [Boolean] True to cancel the event. </param>
Public Property Let Cancel(ByVal a_value As Boolean)
    This.Cancel = a_value
End Property


