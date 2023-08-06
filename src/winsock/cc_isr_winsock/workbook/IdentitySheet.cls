VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdentitySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const PROLOGIX_PORT_NO As Integer = 1234

Private Const m_repeatCountCell As String = "B1"
Private Const m_inputHostCell As String = "B2"
Private Const m_inputPortCell As String = "B3"
Private Const m_receiveTimeoutCell As String = "B4"
Private Const m_errorMessageCell As String = "B5"

Private Const m_socketAddressCell As String = "C2"
Private Const m_connectedSocketIdCell As String = "D2"
Private Const m_connectedCell As String = "E2"
Private Const m_sentMessageCell As String = "F2"
Private Const m_receivedMessageLengthCell As String = "G2"
Private Const m_receivedMessageCell As String = "H2"
Private Const m_statusByteCell As String = "F3"
Private Const m_averageElapsedTimeCell As String = "H3"
Private Const m_elapsedTimeCell As String = "H4"
Private Const m_elapsedTimeFormat As String = "0.0"
Private Const m_identityCell As String = "H2"
Private m_delayStopper As cc_isr_Core.StopWatch

''' <summary>   Reads the identity form the instrument with the
''' IP address and port defined in the parent sheet. </summary>
Private Sub ReadIdentityButton_Click()

    Dim p_host As String: p_host = Range(m_inputHostCell).value
    Dim p_port As Integer: p_port = Range(m_inputPortCell).value
    Dim p_repeatCount As Integer: p_repeatCount = Range(m_repeatCountCell).value
    Dim p_timeout As Integer: p_timeout = Range(m_receiveTimeoutCell).value
    
    Set m_delayStopper = cc_isr_Core.Constructor.CreateStopWatch
    
    If Me.UsingWinsockReadRawButton.value Then
        ' read identity using basic winsock.
        ReadIdentity p_host, p_port, p_repeatCount, p_timeout
    ElseIf Me.UsingTcpClientOptionButton Then
        ' read identity using the Tcp Client.
        TcpClientReadIdentity p_host, p_port, p_repeatCount, p_timeout
    ElseIf Me.UseWinsockWithBufferOptionButton Then
        ' read identity using winsock with buffer.
        ReadIdentityUsingBuffer p_host, p_port, p_repeatCount, p_timeout
    End If
    
End Sub

''' <summary>   Sends the identity query and reads back the identity from the instrument. </summary>
''' <param name="a_host">        [String] The host IPv4 dotted address. </param>
''' <param name="a_port">        [Long] the host port number. </param>
''' <param name="repeatCount">   [Integer] The number of time to repeat the query. </param>
''' <param name="a_timeout">     [Long] receive timeout in milliseconds. </param>
Private Sub ReadIdentity(ByVal a_host As String, ByVal a_port As Long, ByVal a_repeatCount As Integer, _
        ByVal a_timeout As Long)

    On Error GoTo Finally
    
    Range(m_socketAddressCell) = vbNullString
    Range(m_connectedSocketIdCell).value = vbNullString
    Range(m_connectedCell).value = vbNullString
    Range(m_sentMessageCell).value = vbNullString
    Range(m_receivedMessageLengthCell).value = vbNullString
    Range(m_receivedMessageCell).value = vbNullString
    Range(m_elapsedTimeCell).value = vbNullString
    Range(m_errorMessageCell).value = vbNullString

    Dim p_sentCount As Integer
    Dim p_readCount As Integer
    
    Dim p_stopper As StopWatch: Set p_stopper = cc_isr_Core.Constructor.CreateStopWatch
    Dim p_socket As New IPv4StreamSocket
    Dim p_totalMilliseconds As Double
    
    ' report the socket number
    Range(m_connectedSocketIdCell).value = p_socket.SocketId
    
    p_socket.OpenConnection a_host, a_port
    
    ' report the connection state
    Range(m_connectedCell).value = p_socket.Connected
    
    If a_repeatCount > 0 And p_socket.Connected Then
    
        Range(m_socketAddressCell).value = p_socket.SocketAddress
    
        Dim p_command As String
    
        If a_port = PROLOGIX_PORT_NO Then
        
            ' set auto read after write
            ' Prologix GPIB-ETHERNET controller can be configured to automatically address
            ' instruments to talk after sending them a command in order to read their response. The
            ' feature called, Read-After-Write, saves the user from having to issue read commands
            ' repeatedly.
            p_command = "++auto 1"

            ' send the command, which may cause Query Unterminated because we are setting the device to talk
            ' where there is nothing to talk.
            p_sentCount = p_socket.SendMessage(p_command & vbLf)
            m_delayStopper.Wait 5

            ' disables front panel operation of the currently addressed instrument.
            
            p_sentCount = p_socket.SendMessage("++llo" & vbLf)
            m_delayStopper.Wait 5

            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            p_sentCount = p_socket.SendMessage("*CLS; *OPC?" & vbLf)
            p_socket.ReceiveRaw
            m_delayStopper.Wait 5

        End If
        
        p_command = "*IDN?"
        Range(m_sentMessageCell).value = p_command
       
        Dim p_loopCount As Integer
        Do While p_loopCount < a_repeatCount
            
            DoEvents
            
            p_loopCount = p_loopCount + 1
            p_stopper.Restart
            
            ' send the command
            p_sentCount = p_socket.SendMessage(p_command & vbLf)
            m_delayStopper.Wait 5
            
            ' receive the reading
            Dim p_identity As String
            p_identity = p_socket.ReceiveRaw()
    
            p_totalMilliseconds = p_totalMilliseconds + p_stopper.ElapsedMilliseconds
            
            Range(m_receivedMessageLengthCell).value = Len(p_identity)
            Range(m_receivedMessageCell).value = p_identity
    
        Loop
  
        Range(m_averageElapsedTimeCell).value = Format(p_totalMilliseconds / p_loopCount, m_elapsedTimeFormat) + " ms"
        Range(m_elapsedTimeCell).value = Format(p_totalMilliseconds, m_elapsedTimeFormat) + " ms"
    
    ElseIf a_repeatCount <= 0 Then
    
        Range(m_receivedMessageCell).value = "testing connect and disconnect; disconnected."
        
    Else
        Range(m_errorMessageCell).value = "connection failed without reporting an exception"
    End If
    
Finally:

    Dim p_errorMessage As String
    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            p_errorMessage = "Reading identity failed " & cc_isr_Core.UserDefinedErrors.FormattedErr & _
                             " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
        End If
    End If
    
    On Error Resume Next
    
    If Not p_socket Is Nothing Then
    
        If p_socket.Connected Then
        
            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            p_sentCount = p_socket.SendMessage("*CLS; *OPC?" & vbLf)
            p_socket.ReceiveRaw
            m_delayStopper.Wait 5
            
            ' enables front panel operation of the currently addressed instrument.
            
            p_sentCount = p_socket.SendMessage("++loc" & vbLf)
            m_delayStopper.Wait 5
            
            p_socket.CloseConnection
        End If
    
    End If
    
    ' terminate the socket, which should release the Winsock dll.
    Set p_socket = Nothing
    
    If Len(p_errorMessage) > 0 Then
        Range(m_errorMessageCell).value = p_errorMessage
    End If
    
    On Error GoTo 0
    

End Sub

''' <summary>   Sends the identity query and reads back the identity from the instrument. </summary>
''' <param name="a_host">        [String] The host IPv4 dotted address. </param>
''' <param name="a_port">        [Long] the host port number. </param>
''' <param name="repeatCount">   [Integer] The number of time to repeat the query. </param>
''' <param name="a_timeout">     [Long] receive timeout in milliseconds. </param>
Private Sub ReadIdentityUsingBuffer(ByVal a_host As String, ByVal a_port As Long, ByVal a_repeatCount As Integer, _
        ByVal a_timeout As Integer)

    On Error GoTo Finally
    
    Range(m_socketAddressCell) = vbNullString
    Range(m_connectedSocketIdCell).value = vbNullString
    Range(m_connectedCell).value = vbNullString
    Range(m_sentMessageCell).value = vbNullString
    Range(m_receivedMessageLengthCell).value = vbNullString
    Range(m_receivedMessageCell).value = vbNullString
    Range(m_elapsedTimeCell).value = vbNullString
    Range(m_errorMessageCell).value = vbNullString

    Dim p_maximumLength As Integer: p_maximumLength = 1024
    Dim p_buffer As String * 1024
    
    Dim p_sentCount As Integer
    Dim p_readCount As Integer
    
    Dim p_stopper As StopWatch: Set p_stopper = cc_isr_Core.Constructor.CreateStopWatch
    Dim p_socket As New IPv4StreamSocket
    Dim p_totalMilliseconds As Double
    
    ' report the socket number
    Range(m_connectedSocketIdCell).value = p_socket.SocketId
    
    p_socket.OpenConnection a_host, a_port
    
    ' report the connection state
    Range(m_connectedCell).value = p_socket.Connected
    
    If a_repeatCount > 0 And p_socket.Connected Then
    
        ' display the socket address
        Range(m_socketAddressCell).value = p_socket.SocketAddress
        
        Dim p_command As String
    
        If a_port = PROLOGIX_PORT_NO Then
        
            ' set auto read after write
            ' Prologix GPIB-ETHERNET controller can be configured to automatically address
            ' instruments to talk after sending them a command in order to read their response. The
            ' feature called, Read-After-Write, saves the user from having to issue read commands
            ' repeatedly.
            
            p_command = "++auto 1"

            ' send the command, which may cause Query Unterminated because we are setting the device to talk
            ' where there is nothing to talk.
            
            p_sentCount = p_socket.SendMessage(p_command & vbLf)
            m_delayStopper.Wait 5

            ' disables front panel operation of the currently addressed instrument.
            
            p_sentCount = p_socket.SendMessage("++llo" & vbLf)
            m_delayStopper.Wait 5

            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            p_sentCount = p_socket.SendMessage("*CLS; *OPC?" & vbLf)
            p_socket.ReceiveRaw
            m_delayStopper.Wait 5

        End If
        
        p_command = "*IDN?"
        Range(m_sentMessageCell).value = p_command
        
        Dim p_loopCount As Integer
        Do While p_loopCount < a_repeatCount
            
            DoEvents
            
            p_loopCount = p_loopCount + 1
            p_stopper.Restart
            
            ' send the command
            p_socket.SendMessage p_command & vbLf
        
            ' receive the reading
            p_readCount = p_socket.ReceiveTerminatedMessage(p_buffer, p_maximumLength, vbLf)
    
            p_totalMilliseconds = p_totalMilliseconds + p_stopper.ElapsedMilliseconds
            
            Range(m_receivedMessageLengthCell).value = p_readCount
            Range(m_receivedMessageCell).value = p_buffer
    
        Loop
  
        Range(m_averageElapsedTimeCell).value = Format(p_totalMilliseconds / p_loopCount, m_elapsedTimeFormat) + " ms"
        Range(m_elapsedTimeCell).value = Format(p_totalMilliseconds, m_elapsedTimeFormat) + " ms"
    
    ElseIf a_repeatCount <= 0 Then
    
        Range(m_receivedMessageCell).value = "testing connect and disconnect; disconnected."
        
    Else
        Range(m_errorMessageCell).value = "connection failed without reporting an exception"
    End If
    
Finally:

    Dim p_errorMessage As String
    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            p_errorMessage = "Reading identity failed " & cc_isr_Core.UserDefinedErrors.FormattedErr & _
                             " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
        End If
    End If
    
    On Error Resume Next
    
    If Not p_socket Is Nothing Then
    
        If p_socket.Connected Then
        
            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            p_sentCount = p_socket.SendMessage("*CLS; *OPC?" & vbLf)
            p_socket.ReceiveRaw
            m_delayStopper.Wait 5
            
            ' enables front panel operation of the currently addressed instrument.
            
            p_sentCount = p_socket.SendMessage("++loc" & vbLf)
            m_delayStopper.Wait 5
            
            p_socket.CloseConnection
        End If
    
    End If
    
    ' terminate the socket, which should release the Winsock dll.
    Set p_socket = Nothing
    
    If Len(p_errorMessage) > 0 Then
        Range(m_errorMessageCell).value = p_errorMessage
    End If

    On Error GoTo 0

End Sub

''' <summary>   Uses the Tcp Client to send the identity query and reads back the identity
''' from the instrument. </summary>
''' <param name="a_host">        [String] The host IPv4 dotted address. </param>
''' <param name="a_port">        [Long] the host port number. </param>
''' <param name="repeatCount">   [Integer] The number of time to repeat the query. </param>
''' <param name="a_timeout">     [Long] receive timeout in milliseconds. </param>
Private Sub TcpClientReadIdentity(ByVal a_host As String, ByVal a_port As Long, ByVal a_repeatCount As Integer, _
        ByVal a_timeout As Integer)

    On Error GoTo Finally
    
    Range(m_socketAddressCell) = vbNullString
    Range(m_connectedSocketIdCell).value = vbNullString
    Range(m_connectedCell).value = vbNullString
    Range(m_sentMessageCell).value = vbNullString
    Range(m_receivedMessageLengthCell).value = vbNullString
    Range(m_receivedMessageCell).value = vbNullString
    Range(m_elapsedTimeCell).value = vbNullString
    Range(m_errorMessageCell).value = vbNullString

    Dim p_stopper As StopWatch: Set p_stopper = cc_isr_Core.Constructor.CreateStopWatch
    Dim p_totalMilliseconds As Double
    
    Dim p_sentCount As Integer
    Dim p_client As TcpClient
    Set p_client = New TcpClient
    
    p_client.OpenConnection a_host, a_port
    
    ' report the socket number
    Range(m_connectedSocketIdCell).value = p_client.TcpClientSocket.SocketId
   
    ' report the connection state
    Range(m_connectedCell).value = p_client.Connected
    
    If a_repeatCount > 0 And p_client.Connected Then
    
        ' display the socket address.
        Range(m_socketAddressCell).value = p_client.TcpClientSocket.SocketAddress
        
        Dim p_command As String
    
        If a_port = PROLOGIX_PORT_NO Then
        
            ' set auto read after write
            ' Prologix GPIB-ETHERNET controller can be configured to automatically address
            ' instruments to talk after sending them a command in order to read their response. The
            ' feature called, Read-After-Write, saves the user from having to issue read commands
            ' repeatedly.
            
            p_command = "++auto 1"

            ' send the command, which may cause Query Unterminated because we are setting the device to talk
            ' where there is nothing to talk.
            
            p_sentCount = p_client.SendMessage(p_command & vbLf)
            m_delayStopper.Wait 5

            
            ' disables front panel operation of the currently addressed instrument.
            
            p_sentCount = p_client.SendMessage("++llo" & vbLf)
            m_delayStopper.Wait 5

            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            p_sentCount = p_client.SendMessage("*CLS; *OPC?" & vbLf)
            p_client.ReceiveRaw
            m_delayStopper.Wait 5
            

        End If
        
        p_command = "*IDN?"
        Range(m_sentMessageCell).value = p_command
        
        Dim p_loopCount As Integer
        Do While p_loopCount < a_repeatCount
            
            DoEvents
            
            p_loopCount = p_loopCount + 1
            p_stopper.Restart
            
            ' send the command
            p_client.SendMessage p_command & vbLf
            m_delayStopper.Wait 5
        
            ' receive the reading
            Dim p_identity As String
            p_identity = p_client.ReceiveRaw()
    
            p_totalMilliseconds = p_totalMilliseconds + p_stopper.ElapsedMilliseconds
            
            Range(m_receivedMessageLengthCell).value = Len(p_identity)
            Range(m_receivedMessageCell).value = p_identity
    
        Loop
  
        Range(m_averageElapsedTimeCell).value = Format(p_totalMilliseconds / p_loopCount, m_elapsedTimeFormat) + " ms"
        Range(m_elapsedTimeCell).value = Format(p_totalMilliseconds, m_elapsedTimeFormat) + " ms"
    
    ElseIf a_repeatCount <= 0 Then
    
        Range(m_receivedMessageCell).value = "testing connect and disconnect; disconnected."
        
    Else
        Range(m_errorMessageCell).value = "connection failed without reporting an exception"
    End If
    
Finally:

    Dim p_errorMessage As String
    If Not Err Is Nothing Then
        If Err.Number <> 0 Then
            p_errorMessage = "Reading identity failed using Tcp Client " & cc_isr_Core.UserDefinedErrors.FormattedErr & _
                             " Last DLL error = 0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
        End If
    End If
    
    On Error Resume Next
    
    If Not p_client Is Nothing Then
        If p_client.Connected Then
            
            ' clear errors if any so as to leave the instrument without errors.
            ' here we add *OPC? to prevent the query unterminated error.
            p_sentCount = p_client.SendMessage("*CLS; *OPC?" & vbLf)
            p_client.ReceiveRaw
            m_delayStopper.Wait 5
            
            ' enables front panel operation of the currently addressed instrument.
            
            p_sentCount = p_client.SendMessage("++loc" & vbLf)
            m_delayStopper.Wait 5

            
            p_client.CloseConnection
        End If
    End If
    
    ' terminate the client
    Set p_client = Nothing
    
    If Len(p_errorMessage) > 0 Then
        Range(m_errorMessageCell).value = p_errorMessage
    End If

    On Error GoTo 0

End Sub


