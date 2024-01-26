VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Net"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Static Net commands."
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Net commands. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Determines the round-trip time (RTT) and hop count to the specified destination. </summary>
''' <remarks>
''' <see href="https://www.fortypoundhead.com/showcontent.asp?artid=24045"/>
''' <see href="http://vbnet.mvps.org/"/>
''' </remarks>
''' <param name="a_ipv4Address">   [in, String] IP address of the destination for which to determine the
'''                                RTT and hop count. </param>
''' <param name="a_maxHops">       [in, Long, 20] Maximum number of hops to search for the destination. If the number
'''                                of hops to the destination exceeds this number, the function terminates the
'''                                search and returns FALSE. </param>
''' <returns>   [Boolean]. If the function succeeds, the return value is TRUE; otherwise, if the function fails,
'''             the return value is FALSE. Call GetLastError to obtain the error code for the failure. <returns>
Public Function Ping(ByVal a_ipv4Address As String, Optional ByVal a_maxHops As Long = 20) As Boolean

    Dim p_hopsCount      As Long
    Dim p_roundTripTime  As Long
    Ping = WinsockApi.GetRoundTripAndHopCount(a_ipv4Address, p_hopsCount, a_maxHops, p_roundTripTime)
    
End Function

''' <summary>   Attempts to connect to the specified host at the specified port. </summary>
''' <param name="a_host">        [String] The host IPv4 dotted address. </param>
''' <param name="a_port">        [Long] the host port number. </param>
''' <param name="a_details">     [String] details the failure. </param>
''' <returns>   [Boolean] True on success; otherwise, false. </returns>
Public Function PingPort(ByVal a_host As String, ByVal a_port As Long, ByRef a_details As String) As Boolean

    Const p_procedureName As String = "PingPort"
    
    On Error GoTo Finally
    
    Dim p_socket As New IPv4StreamSocket
    
    If p_socket.SocketId <> 0 Then
    
        If p_socket.OpenConnection(a_host, a_port) Then
        
            PingPort = True
        
        Else
        
            a_details = "Failed connecting to " & a_host & ":" & VBA.CStr(a_port) & "."
            PingPort = False
            
        End If
    
    Else
    
        a_details = "Failed initializing a socket. Socket id set to " & VBA.CStr(p_socket.SocketId) & "."
        PingPort = False
    
    End If
    
Finally:

    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            PingPort = False
            cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
            a_details = "Error connecting to " & a_host & ":" & VBA.CStr(a_port) & ". " & _
                        cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        End If
    End If
    
    On Error Resume Next
    
    If Not p_socket Is Nothing Then
    
        If p_socket.Connected Then
            
            cc_isr_Core_IO.Factory.NewStopwatch.Wait 5
            
            p_socket.CloseConnection
        
        End If
    
    End If
    
    ' terminate the socket, which should release the Winsock dll.
    Set p_socket = Nothing
    
    On Error GoTo 0

End Function

''' <summary>   Attempts to connect to the Prologix device as '192.168.0.252:1234'. </summary>
''' <param name="a_host">   [String, 192.168.0.252] The host IPv4 dotted address. </param>
''' <param name="a_port">   [Long, 1234] the host port number. </param>
''' <returns>   [Boolean] True on success; otherwise, false and output a debug message. </returns>
Public Function PingPrologix(Optional ByVal a_host As String = "192.168.0.252", _
        Optional ByVal a_port As Long = 1234) As Boolean
    
    Dim p_details As String
    
    PingPrologix = Me.PingPort(a_host, a_port, p_details)
    
    If Not PingPrologix Then _
        Debug.Print "@" & ThisWorkbook.VBProject.Name & "." & "Net.PingPrologix. " & p_details
   
End Function


