Attribute VB_Name = "SocketQueryIdentityTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Socket query identity Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   This class properties. </summary>
Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
    Host As String
    Port As Long
    PrologixPort As Long
    SocketReceiveTimeout As Integer
    Socket As IPv4StreamSocket
    DelayStopper As cc_isr_Core_IO.Stopwatch
    ErrTracer As IErrTracer
End Type

Private This As this_

''' <summary>   Runs the specified test. </summary>
Public Sub RunTest(ByVal a_testNumber As Integer)
    BeforeEach
    Select Case a_testNumber
        Case 1
            TestSocketShouldRawQueryIdentity
        Case 2
            TestSocketShouldBufferQueryIdentity
        Case Else
    End Select
    AfterEach
End Sub

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 2
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    BeforeAll
    Dim p_testNumber As Integer
    For p_testNumber = 1 To 1
        RunTest p_testNumber
        DoEvents
    Next p_testNumber
    AfterAll
End Sub

''' <summary>   Prepares all tests. </summary>
Public Sub BeforeAll()

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    This.Name = "SocketQueryIdentityTests"
    
    This.TestNumber = 0
    This.Host = "192.168.0.252"
    This.Port = 1234
    This.PrologixPort = 1234
    This.SocketReceiveTimeout = 100
    
    This.TestNumber = 0
    
    Set This.ErrTracer = New ErrTracer
    
    Set This.BeforeAllAssert = Assert.IsTrue(True, "initialize the overall assert.")
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    Set This.DelayStopper = cc_isr_Core_IO.Factory.NewStopwatch
        
    Set This.Socket = cc_isr_Winsock.Factory.NewIPv4StreamSocket()
   
    This.Socket.OpenConnection This.Host, This.Port
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        Set This.BeforeAllAssert = Assert.Inconclusive("Failed preparing all tests: " & _
            p_leftoverErrorMessage)
        This.ErrTracer.TraceError p_leftoverErrorMessage
    
    ElseIf This.Socket.Connected Then
        
        Set This.BeforeAllAssert = Assert.Pass("Connected")
    
    Else
        
        Set This.BeforeAllAssert = Assert.Inconclusive("IPV4 Stream Socket should be connected")
    
    End If

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Prepares each test before it is run. </summary>
Public Sub BeforeEach()

    Const p_procedureName As String = "BeforeEach"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
        
        Set This.BeforeEachAssert = IIf(This.Socket.Connected, _
            Assert.IsTrue(True, "Connected"), _
            Assert.Inconclusive("IPV4 Stream Socket should be connected"))
    
    Else
    
        Set This.BeforeEachAssert = Assert.Inconclusive(This.BeforeAllAssert.AssertMessage)
    
    End If
    
    This.TestNumber = This.TestNumber + 1
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    If This.BeforeEachAssert.AssertSuccessful Then
    
        Set This.BeforeEachAssert = Assert.AreEqual(0, Err.Number, _
            "Error Number should be 0.")
            
    End If
    
    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_reply As String

    ' prime the Prologix device
    If This.Port = This.PrologixPort Then
    
        ' set auto read after write
        ' Prologix GPIB-ETHERNET controller can be configured to automatically address
        ' instruments to talk after sending them a command in order to read their response. The
        ' feature called, Read-After-Write, saves the user from having to issue read commands
        ' repeatedly.
        p_command = "++auto 1"

        ' send the command, which may cause Query Unterminated because we are setting the device to talk
        ' where there is nothing to talk.
        p_sentCount = This.Socket.SendMessage(p_command & VBA.vbLf)
        This.DelayStopper.Wait 5

        ' disables front panel operation of the currently addressed instrument.
        
        p_sentCount = This.Socket.SendMessage("++llo" & VBA.vbLf)
        This.DelayStopper.Wait 5

    End If

    ' clear execution state before each test.
    ' clear errors if any so as to leave the instrument without errors.
    ' here we add *OPC? to prevent the query unterminated error.

    p_sentCount = This.Socket.SendMessage("*CLS;*WAI;*OPC?" & VBA.vbLf)
    This.DelayStopper.Wait 5
    p_reply = This.Socket.ReceiveRaw
    This.DelayStopper.Wait 5
    
    Set This.BeforeEachAssert = Assert.AreEqual("1", p_reply, _
            "Operation completion should send the correct reply.")
                    
    This.TestNumber = This.TestNumber + 1

    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    If This.BeforeEachAssert.AssertSuccessful Then
    
        Set This.BeforeEachAssert = Assert.AreEqual(0, Err.Number, _
            "Error Number should be 0.")
            
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        Set This.BeforeAllAssert = Assert.Inconclusive("Failed preparing test #" & VBA.CStr(This.TestNumber) & ": " & _
            p_leftoverErrorMessage)
        This.ErrTracer.TraceError p_leftoverErrorMessage
    
    End If

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases test elements after each tests is run. </summary>
Public Sub AfterEach()
    
    Const p_procedureName As String = "AfterEach"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler


    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_reply As String

    If This.BeforeEachAssert.AssertSuccessful Then
    
        ' clear errors if any so as to leave the instrument without errors.
        p_sentCount = This.Socket.SendMessage("*CLS;*WAI;*OPC?" & VBA.vbLf)
        This.DelayStopper.Wait 5
        p_reply = This.Socket.ReceiveRaw
        This.DelayStopper.Wait 5

        ' Restore Prologix device
        If This.BeforeEachAssert.AssertSuccessful And This.Port = This.PrologixPort Then
        
            p_command = "++auto 0"

            ' send the command, which may cause Query Unterminated because we are setting the device to talk
            ' where there is nothing to talk.
            p_sentCount = This.Socket.SendMessage(p_command & VBA.vbLf)
            This.DelayStopper.Wait 5

            ' restore front panel operation of the currently addressed instrument.
            
            p_sentCount = This.Socket.SendMessage("++loc" & VBA.vbLf)
            This.DelayStopper.Wait 5

        End If

    End If
    
    Set This.BeforeEachAssert = Nothing
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        This.ErrTracer.TraceError "Error(s) were stacked unwinding test #" & _
            VBA.CStr(This.TestNumber) & ": " & p_leftoverErrorMessage
    
    End If

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases the test class after all tests run. </summary>
Public Sub AfterAll()
    
    Const p_procedureName As String = "AfterAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' disconnect if connected
    If Not This.Socket Is Nothing Then _
        This.Socket.CloseConnection

    Set This.Socket = Nothing

    Set This.BeforeAllAssert = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        This.ErrTracer.TraceError "Errors were stacked unwinding all tests: " & p_leftoverErrorMessage
    
    End If

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Unit test. Asserts that the stream socket should query a device identity. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestSocketShouldRawQueryIdentity() As Assert

    Const p_procedureName As String = "TestSocketShouldRawQueryIdentity"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String: p_command = "*IDN?"
    Dim p_sentCount As Integer
    Dim p_identity As String
    
    If p_outcome.AssertSuccessful Then
            
        ' send the command
        p_sentCount = This.Socket.SendMessage(p_command & VBA.vbLf)
        This.DelayStopper.Wait 5
            
        p_identity = This.Socket.ReceiveRaw()

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestSocketShouldRawQueryIdentity")
    
    Set TestSocketShouldRawQueryIdentity = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts that the stream socket should query a device identity. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestSocketShouldBufferQueryIdentity() As Assert

    Const p_procedureName As String = "TestSocketShouldBufferQueryIdentity"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String: p_command = "*IDN?"
    Dim p_sentCount As Integer
    Dim p_identity As String
    Dim p_maximumLength As Integer: p_maximumLength = 1024
    Dim p_buffer As String * 1024
    Dim p_readCount As Integer
    
    If p_outcome.AssertSuccessful Then
            
        ' send the command
        p_sentCount = This.Socket.SendMessage(p_command & VBA.vbLf)
        This.DelayStopper.Wait 5
    
        ' receive the reading
        p_readCount = This.Socket.ReceiveTerminatedMessage(p_buffer, p_maximumLength, VBA.vbLf)
    
        p_identity = p_buffer

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestSocketShouldBufferQueryIdentity")
    
    Set TestSocketShouldBufferQueryIdentity = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function



