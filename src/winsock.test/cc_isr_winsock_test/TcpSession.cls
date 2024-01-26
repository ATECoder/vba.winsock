VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A TCP Session capable of communicating with a GPIB-Lan controller. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Socket As IPv4StreamSocket
    ReadAfterWriteDelay As Long
    GpibLanControllerPort As Long
    Termination As String
    DelayStopper As cc_isr_Core_IO.Stopwatch
    AppendTermination As AppendTerminationOption
    AutoAssertTalk As Boolean
    EndOrIdentify As Boolean
End Type

Private This As this_

Private WithEvents m_Socket As IPv4StreamSocket
Attribute m_Socket.VB_VarHelpID = -1

''' <summary>   GPIB Lan Controller Termination options. </summary>
Public Enum AppendTerminationOption

    ''' <summary>   Enum option for appending carriage return and line feed. </summary>
    CarriageReturnLineFeed = 0
    ''' <summary>   Enum option for appending carriage return. </summary>
    CarriageReturn = 1
    ''' <summary>   Enum option for appending line feed. </summary>
    LineFeed = 2
    ''' <summary>   Enum option for appending nothing. </summary>
    AppendNothing = 3

End Enum


''' <summary>   Handles the Class Initialize event. </summary>
Private Sub Class_Initialize()
    This.GpibLanControllerPort = 1234
    This.ReadAfterWriteDelay = 1
    This.Termination = VBA.vbLf
    Set This.DelayStopper = cc_isr_Core_IO.Factory.NewStopwatch
End Sub

''' <summary>   Handles the Class terminate event; disposes of the <see cref="Socket"/>. </summary>
Private Sub Class_Terminate()
    Set This.Socket = Nothing
    Set m_Socket = Nothing
    Set This.DelayStopper = Nothing
End Sub

''' <summary>   Returns this unconnected Tcp Client instance. No actual initialization is done
''' at this time. </summary>
''' <remarks>   Raises an invalid operation exception is a socket is already connected. </remarks>
''' <param name="a_socket ">   [IPv4StreamSocket] An INET stream socket. </param>
''' <returns>   [<see cref="TcpSession"/>]. </returns>
Public Function Initialize(ByVal a_socket As IPv4StreamSocket) As TcpSession
    
    Set This.Socket = a_socket
    Set m_Socket = a_socket
    Set Initialize = Me
    
End Function

''' <summary>   Gets the Tcp Client socket. </summary>
''' <value>   [<see cref="IPv4StreamSocket"/>]. </value>
Public Property Get Socket() As IPv4StreamSocket
    Set Socket = This.Socket
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  I/O
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the termination to append to output messages. </summary>
''' <value>   [String]. </value>
Public Property Get Termination() As String
    Termination = This.Termination
End Property

''' <summary>   Sets the termination to append to output messages. </summary>
''' <param name="a_value">   [String] the termination. </param>
Public Property Let Termination(ByVal a_value As String)
    This.Termination = a_value
End Property

''' <summary>   Gets the Read After Write Delay in milliseconds. </summary>
''' <value>   [Integer]. </value>
Public Property Get ReadAfterWriteDelay() As Integer
    ReadAfterWriteDelay = This.ReadAfterWriteDelay
End Property

''' <summary>   Sets the Read After Write Delay Delay in milliseconds. </summary>
''' <param name="a_value">   [Integer] the Read After Write Delay Delay in milliseconds. </param>
Public Property Let ReadAfterWriteDelay(ByVal a_value As Integer)
    This.ReadAfterWriteDelay = a_value
End Property

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Device send and receive
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Sends an unterminate message to the server. </summary>
''' <param name="a_message">   [String] The unterminated message. </param>
''' <returns>   [Long] The number of bytes that were sent to the server. </returns>
Public Function SendMessage(ByVal a_message As String) As Long
    SendMessage = This.Socket.SendMessage(a_message & This.Termination)
    This.DelayStopper.Wait This.ReadAfterWriteDelay
End Function

''' <summary>   Receives a message from the server until reaching the specified termination. </summary>
''' <remarks>   The <see cref="ReceiveTimeout"/> must be set long enough the exceed any instrument
''' or controller delays. </remarks>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] detasils the failure reason. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [Integer] If no error, the number of bytes that were received.
''' Otherwise, RECEIVE_ERROR (-1).
''' </returns>
Public Function TryReceive(ByRef a_buffer As String, _
                           ByRef a_details As String, _
                           Optional ByVal a_trimEnd As Boolean = True, _
                           Optional ByVal a_termination As String = vbLf) As Long
                           
    If Me.GpibLanControllerAttached Then
        Me.SendMessage "++read eoi"
    End If
                           
    TryReceive = This.Socket.TryReceive(a_buffer, a_details, a_trimEnd, a_termination)
                        
End Function

''' <summary>   Sends an unterminated query message and reads back from the server until reaching the specified
''' termination. </summary>
''' <param name="a_message">       [String] The unterminated query message. </param>
''' <param name="a_buffer">        [Out, String] the received message. </param>
''' <param name="a_details">       [Out, String] detasils the failure reason. </param>
''' <param name="a_trimEnd">       [Optional, Boolean, True] True to return the string without
'''                                the termination. </param>
''' <param name="a_termination">   [Optional, String, Line Feed] The character that signifies the
'''                                end of the stream. </param>
''' <returns>   [Integer] If no error, the number of bytes that were received.
''' Otherwise, RECEIVE_ERROR (-1).
''' </returns>
Public Function TryQuery(ByVal a_message As String, _
                         ByRef a_buffer As String, _
                         ByRef a_details As String, _
                         Optional ByVal a_trimEnd As Boolean = True, _
                         Optional ByVal a_termination As String = VBA.vbLf) As Long
    If Me.SendMessage(a_message) > 0 Then
        TryQuery = Me.TryReceive(a_buffer, a_details, a_trimEnd, a_termination)
    Else
        a_details = "Query command is empty."
        TryQuery = -1
    End If
  
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  GPIB Lan Controller Methods
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Gets the port number of the GPIB Lan controller. </summary>
''' <value>   [Long]. </value>
Public Property Get GpibLanControllerPort() As Long
    GpibLanControllerPort = This.GpibLanControllerPort
End Property

''' <summary>   Sets the port number of the GPIB Lan controller. </summary>
''' <param name="a_value">   [Long] the GPIB Lan controller port. </param>
Public Property Let GpibLanControllerPort(ByVal a_value As Long)
    This.GpibLanControllerPort = a_value
End Property

''' <summary>   True if a GPIB Lan controller is attached. </summary>
''' <value>   [Boolean]. </value>
Public Property Get GpibLanControllerAttached() As Boolean
    If This.Socket Is Nothing Then
        GpibLanControllerAttached = False
    Else
        GpibLanControllerAttached = This.Socket.Port = This.GpibLanControllerPort
    End If
End Property

''' <summary>   Tries to get and validate a value returning the actual value and true if
''' the correct value was set or false if not. </summary>
''' <remarks>   This method assumes that the bare command is the query command.  </remarks>
''' <param name="a_command">   [String] the bare command, e.g.,<c>++eos</c>. </param>
''' <param name="a_value">     [String] the expected value to get. </param>
''' <param name="a_result">    [Out, String] the actual value that was set. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <value>   [Boolean]. </value>
Public Function TryGetValue(ByVal a_command As String, ByVal a_value As String, _
    ByRef a_result As String, ByRef a_details As String) As Boolean

    ' send the query command.
    Dim p_sentCount As Integer
    p_sentCount = Me.SendMessage(a_command)
    
    Dim p_receiveCount As Integer
    p_receiveCount = This.Socket.TryReceive(a_result, a_details)
    
    ' the receive count is negative if error
    If 0 > p_receiveCount Then
        a_details = "Command '" & a_command & "' failed; " & a_details
        TryGetValue = False
    Else
        If 1 = VBA.InStr(1, a_value, a_result, vbTextCompare) Then
            TryGetValue = True
        Else
            a_details = "Command '" & a_command & "' failed; Result '" & a_result & "' not equals expected value '" & a_value & "'."
            TryGetValue = False
        End If
    End If

End Function


' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  GPIB Lan Commands
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++auto
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <sumamry>   Automatically address instruments to talk after sending a command. </summary>
''' <remarks>
''' The GPIB-Lan controller can be configured to automatically address
''' instruments to talk after sending them a command in order to read their response. The
''' feature called, Read-After-Write, saves the user from having to issue read commands
''' repeatedly. This command enabled or disabled the Read-After-Write feature.
''' In addition, auto command also addresses the instrument at the currently specified
''' address to TALK or LISTEN. <c>++auto 0</c> addresses the instrument to LISTEN and
''' <c>++auto 1</c> addresses the instrument to TALK.
''' If the command is issued without any arguments it returns the current state of the
''' readafter-write feature.
''' SYNTAX: <c>++auto [0|1]</c>
''' Some instruments generate “Query Unterminated” or “-420” error if they are addressed
''' to talk after sending a command that does not generate a response (often called nonquery commands).
''' In effect the instrument is saying, I have been asked to talk but I have
''' nothing to say. The error is often benign and may be ignored. Otherwise, use the
''' <c>++read</c> command to read the instrument response.
''' </remarks>
''' <value>   [String]. </value>
Public Property Get AutoAssertTalkCommand() As String
    AutoAssertTalkCommand = "++auto"
End Property

''' <summary>   GPIB Lan Controller Auto Assert Talk. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AutoAssertTalk() As Boolean
    AutoAssertTalk = This.AutoAssertTalk
End Property

''' <summary>   Gets the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <returns>   [Boolean]. </returns>
Public Function AutoAssertTalkGetter() As Boolean

    Dim p_sentCount As Integer
    p_sentCount = Me.SendMessage(Me.AutoAssertTalkCommand)
    
    Dim p_result As String
    Dim p_details As String
    Dim p_receiveCount As Integer
    p_receiveCount = This.Socket.TryReceive(p_result, p_details)
    
    If p_receiveCount > 0 Then
        If Not Me.TryParseAutoAssertTalk(p_result, p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".AutoAssertTalkGetter", _
                "The return value '" & p_result & "' is invalid for the '" & Me.AutoAssertTalkCommand & "' command."
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".AutoAssertTalkGetter", p_details
    End If
    AutoAssertTalkGetter = This.AutoAssertTalk

End Function

''' <summary>   Tries to parse the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryParseAutoAssertTalk(ByVal a_value As String, ByRef a_details As String) As Boolean
    
    If "1" = a_value Then
        This.AutoAssertTalk = True
    ElseIf "0" = a_value Then
        This.AutoAssertTalk = False
    Else
        a_details = "The return value '" & a_value & "' is invalid for the '" & Me.AutoAssertTalkCommand & "' command."
    End If
    
End Function

''' <summary>   Sets the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
Public Sub AutoAssertTalkSetter(ByVal a_value As Boolean)

    Dim p_details As String
    If Not Me.TryAutoAssertTalkSetter(a_value, p_details) Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".AutoAssertTalkSetter", p_details

End Sub

''' <summary>   Tries to set the GPIB Lan Controller Auto Assert Talk state. </summary>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryAutoAssertTalkSetter(ByVal a_value As Boolean, ByRef a_details As String) As Boolean
    
    Dim p_command As String
    Dim p_value As String
    Dim p_result As String
    Dim p_details As String
    Dim p_sentCount As Long
    
    p_command = Me.AutoAssertTalkCommand
    p_value = VBA.CStr(a_value)
    
    ' send the command
    p_sentCount = Me.SendMessage(p_command & " " & p_value)
    
    ' validate the command
    If Me.TryGetValue(p_command, p_value, p_result, p_details) Then
    
        TryAutoAssertTalkSetter = Me.TryParseAutoAssertTalk(p_result, p_details)
    
    Else
    
        TryAutoAssertTalkSetter = False
    
    End If
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++eos
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   The Append Termination command or EOS, specified the GPIB Lan Controller
''' option for appending a termination string to the sent message. </summary>
''' <remarks>  This command specifies GPIB termination characters. When data from host is received
''' over the network, all non-escaped LF, CR and ESC characters are removed and GPIB
''' terminators, as specified by this command, are appended before sending the data to
''' instruments.
'''
''' Although the manual indicates that "This command does not affect data from instruments
''' received over GPIB port.", we found that messages received from the instrument were
''' terminated with CR LF when EOS was set at 0 and LF when EOS was set at 3.
''' SYNTAX: <c>++eos [0|1|2|3]
''' </remarks>
''' <value>   [String]. </value>
Public Property Get AppendTerminationCommand() As String
    AppendTerminationCommand = "++eos"
End Property

''' <summary>   GPIB Lan Controller Termination options. </summary>
''' <value>   [<see cref="AppendTerminationOption"/>]. </value>
Public Property Get AppendTermination() As AppendTerminationOption
    AppendTermination = This.AppendTermination
End Property

''' <summary>   Gets the GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <returns>   [Boolean]. </returns>
Public Function AppendTerminationGetter() As AppendTerminationOption

    Dim p_sentCount As Integer
    p_sentCount = Me.SendMessage(Me.AppendTerminationCommand)
    
    Dim p_result As String
    Dim p_details As String
    Dim p_receiveCount As Integer
    p_receiveCount = This.Socket.TryReceive(p_result, p_details)
    
    If p_receiveCount > 0 Then
        
        If Not Me.TryParseAppendTermination(p_result, p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".AppendTerminationGetter", _
                "The return value '" & p_result & "' is invalid for the '" & Me.AppendTerminationCommand & "' command."
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".AppendTerminationGetter", p_details
    End If
    AppendTerminationGetter = This.AppendTermination

End Function

''' <summary>   Sets the GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_value">     [<see cref="AppendTerminationOption"/>] the option to set. </param>
Public Sub AppendTerminationSetter(ByVal a_value As AppendTerminationOption)

    Dim p_details As String
    If Not Me.TryAppendTerminationSetter(a_value, p_details) Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".AppendTerminationSetter", p_details

End Sub

''' <summary>   Tries to parse GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [<see cref="AppendTerminationOption"/>]. </returns>
Public Function TryParseAppendTermination(ByVal a_value As String, ByRef a_details As String) As Boolean
    
    Dim p_value As Integer
    If cc_isr_core.StringExtensions.TryParseInteger(a_value, p_value, a_details) Then
        This.AppendTermination = p_value
        TryParseAppendTermination = True
    Else
        TryParseAppendTermination = False
    End If
    
End Function

''' <summary>   Tries to set the GPIB Lan Controller <see cref="AppendTerminationOption"/>. </summary>
''' <param name="a_value">     [<see cref="AppendTerminationOption"/>] The value to set. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryAppendTerminationSetter(ByVal a_value As AppendTerminationOption, _
    ByRef a_details As String) As Boolean
    
    Dim p_command As String
    Dim p_value As String
    Dim p_result As String
    Dim p_details As String
    Dim p_sentCount As Long
    
    p_command = Me.AppendTerminationCommand
    p_value = VBA.CStr(a_value)
    
    ' send the command
    p_sentCount = Me.SendMessage(p_command & " " & p_value)
    
    ' validate the command
    If Me.TryGetValue(p_command, p_value, p_result, p_details) Then
    
        TryAppendTerminationSetter = Me.TryParseAppendTermination(p_result, p_details)
    
    Else
    
        TryAppendTerminationSetter = False
    End If
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++eoi
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <sumamry>   This command enables or disables the assertion of the EOI signal with the
''' last character of any command sent over GPIB port. </summary>
''' <remarks>   Some instruments require EOI signal to be asserted in order to properly
''' detect the end of a command. SYNTAX: <c>++eoi [0|1]</c>. </remarks>
Public Property Get EndOrIdentifyCommand() As String
    EndOrIdentifyCommand = "++eoi"
End Property

''' <summary>   GPIB Lan Controller End Or Identify (EOI). </summary>
''' <value>   [Boolean]. </value>
Public Property Get EndOrIdentify() As Boolean
    EndOrIdentify = This.EndOrIdentify
End Property

''' <summary>   Gets the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <returns>   [Boolean]. </returns>
Public Function EndOrIdentifyGetter() As Boolean

    Dim p_sentCount As Integer
    p_sentCount = Me.SendMessage(Me.EndOrIdentifyCommand)
    
    Dim p_result As String
    Dim p_details As String
    Dim p_receiveCount As Integer
    p_receiveCount = This.Socket.TryReceive(p_result, p_details)
    
    If p_receiveCount > 0 Then
        If Not TryParseEndOrIdentify(p_result, p_details) Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".EndOrIdentifyGetter", _
                "The return value '" & p_result & "' is invalid for the '" & Me.EndOrIdentifyCommand & "' command."
    Else
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".EndOrIdentifyGetter", p_details
    End If
    EndOrIdentifyGetter = This.EndOrIdentify

End Function

''' <summary>   Tries to parse the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <param name="a_value">     [String] The value to parse. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryParseEndOrIdentify(ByVal a_value As String, ByRef a_details As String) As Boolean
    
    If "1" = a_value Then
        This.EndOrIdentify = True
    ElseIf "0" = a_value Then
        This.EndOrIdentify = False
    Else
        a_details = "The return value '" & a_value & "' is invalid for the '" & Me.EndOrIdentifyCommand & "' command."
    End If
    
End Function

''' <summary>   Sets the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
Public Sub EndOrIdentifySetter(ByVal a_value As Boolean)

    Dim p_details As String
    If Not Me.TryEndOrIdentifySetter(a_value, p_details) Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".EndOrIdentifySetter", p_details

End Sub

''' <summary>   Tries to set the GPIB Lan Controller End Or Identify (EOI) state. </summary>
''' <param name="a_value">     [Boolean] True to set or false to turn off. </param>
''' <param name="a_details">   [Out, String] details the failure reason. </param>
''' <returns>   [Boolean]. </returns>
Public Function TryEndOrIdentifySetter(ByVal a_value As Boolean, ByRef a_details As String) As Boolean
    
    Dim p_command As String
    Dim p_value As String
    Dim p_result As String
    Dim p_details As String
    Dim p_sentCount As Long
    
    p_command = Me.EndOrIdentifyCommand
    p_value = VBA.CStr(a_value)
    
    ' send the command
    p_sentCount = Me.SendMessage(p_command & " " & p_value)
    
    ' validate the command
    If Me.TryGetValue(p_command, p_value, p_result, p_details) Then
    
        TryEndOrIdentifySetter = Me.TryParseEndOrIdentify(p_result, p_details)
    
    Else
    
        TryEndOrIdentifySetter = False
    
    End If
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  ++spoll
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' summary>   Reads the status byte of the current GPIB instrument. </summary>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <returns>   [Integer] The status byte or RECEIVE_ERROR (-1) if failed receiving a reading or
''' failed parsing the reading to an integer.
''' </returns>
Public Function SerialPoll(ByRef a_details As String) As Integer
    
    Dim p_command As String
    p_command = "++spoll"
    
    Dim p_sentCount As Integer
    p_sentCount = Me.SendMessage(p_command)
    
    Dim p_reply As String
    Dim p_receiveCount As Integer
    p_receiveCount = This.Socket.TryReceive(p_reply, a_details)
    Dim p_statusByte As Integer
    If 0 < p_receiveCount Then
        If Not cc_isr_core.StringExtensions.TryParseInteger(p_reply, p_statusByte, a_details) Then
            p_statusByte = -1
        End If
    Else
        p_statusByte = -1
    End If
    SerialPoll = p_statusByte

End Function

''' summary>   Waits for the expected status bits or timeout. </summary>
''' <param name="a_bitsStatus"/>   [Integer] The expected status of the specified status bits. </param>
''' <param name="a_statusBits"/>   [Integer] The expected status bits. </param>
''' <param name="a_timeout"/>      [Long] The timeout in milliseconds. </param>
''' <param name="a_statusByte"/>   [Out, Integer] The last received status byte.
'''                                Negative if serial poll failed. </param>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <returns>   [Boolean] True if the status byte has the expected bits value. </returns>
Public Function AwaitStatusBits(ByVal a_bitsStatus As Integer, ByVal a_statusBits As Integer, _
    ByVal a_timeout As Long, ByRef a_statusByte As Integer, ByRef a_details As String) As Boolean

    Dim p_gotIt As Boolean
    Dim p_stopper As cc_isr_Core_IO.Stopwatch
    Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
    p_stopper.Restart
    
    Do
        DoEvents
        ' read the status bit.
        a_statusByte = Me.SerialPoll(a_details)
        p_gotIt = (a_statusByte >= 0) And (a_bitsStatus = (a_statusBits And a_statusByte))
    Loop Until p_gotIt Or (a_statusByte < 0) Or (p_stopper.ElapsedMilliseconds > a_timeout)
    
    AwaitStatusBits = p_gotIt
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'
'  Socket Event Handlers
'
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Handles the <see cref="cc_isr_Winsock.IPv4SreamSocket.ConnectionChanged"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangedEventArgs"/> event arguments. </param>
Private Sub m_Socket_ConnectionChanged(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangedEventArgs)

    Const p_procedureName As String = "m_Socket_ConnectionChanged"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_eventArgs.Connected Then
    
        If This.Socket.Port = This.GpibLanControllerPort Then
        
            ' prime the GPIB-Lan controller
            '
            ' EOS and EOI were set per these recommendations:
            '
            ' https://groups.io/g/HP-Agilent-Keysight-equipment/topic/86224398
            
            Dim p_command As String
            Dim p_value As String
            Dim p_result As String
            Dim p_details As String
            Dim p_sentCount As Long
            
            ' set the GPIB termination characters to none - do not append termination characters.
            
            p_command = "++eos"
            p_value = "3"
            
            ' send the command
            p_sentCount = Me.SendMessage(p_command & " " & p_value)
            
            ' validate the command
            If Not Me.TryGetValue(p_command, p_value, p_result, p_details) Then _
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
            
            p_command = "++eoi"
            p_value = "1"
            
            ' send the command
            p_sentCount = Me.SendMessage(p_command & " " & p_value)
            
            ' validate the command
            If Not Me.TryGetValue(p_command, p_value, p_result, p_details) Then _
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
            
            p_command = "++auto"
            p_value = "0"
            
            ' send the command
            p_sentCount = Me.SendMessage(p_command & " " & p_value)
            
            ' validate the command
            If Not Me.TryGetValue(p_command, p_value, p_result, p_details) Then _
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
            
            ' disable front panel operation of the currently addressed instrument.
            p_sentCount = Me.SendMessage("++llo")
        
        End If
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the <see cref="cc_isr_Winsock.IPv4SreamSocket.ConnectionChanging"/> event. </summary>
''' <param name="a_eventArgs">   Reference to the <see cref="cc_isr_Winsock.ConnectionChangingEventArgs"/> event arguments. </param>
Private Sub m_Socket_ConnectionChanging(ByVal a_eventArgs As cc_isr_Winsock.ConnectionChangingEventArgs)

    Const p_procedureName As String = "m_Socket_ConnectionChanging"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_eventArgs.Connected Then
    
        If Me.GpibLanControllerAttached Then
        
            Dim p_command As String
            Dim p_value As String
            p_command = "++auto"
            p_value = "0"
            
            ' send the command
            Me.SendMessage p_command & " " & p_value
            
            ' validate the command
            Dim p_result As String
            Dim p_details As String
            If Not Me.TryGetValue(p_command, p_value, p_result, p_details) Then _
                cc_isr_Core_IO.UserDefinedErrors.EnqueueError cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, p_details
        
            ' restore front panel operation of the currently addressed instrument.
            Me.SendMessage "++loc"
    
        End If
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, TypeName(Me), ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub



