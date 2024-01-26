Attribute VB_Name = "SocketSerialPollQueryTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Socket Serial Poll query identity Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   This class properties. </summary>
Private Type this_
    Name As String
    TestNumber As Integer
    PreviousTestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    Address As String
    PrologixPort As Long
    Socket As IPv4StreamSocket
    Termination As String
    ReceiveTimeout As Long
    ReadAfterWriteDelay As Integer
    AssertTalkOnWrite As Boolean
    DelayStopper As cc_isr_Core_IO.Stopwatch
    TestStopper As cc_isr_Core_IO.Stopwatch
    ErrTracer As IErrTracer
    IdentityCompany As String
    TestCount As Integer
    RunCount As Integer
    PassedCount As Integer
    FailedCount As Integer
    InconclusiveCount As Integer
End Type

Private This As this_

''' <summary>   Runs the specified test. </summary>
Public Function RunTest(ByVal a_testNumber As Integer) As cc_isr_Test_Fx.Assert
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.TestNumber = a_testNumber
    BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestSocketShouldConnect
        Case 2
            Set p_outcome = TestSocketShouldQueryIdentity
        Case 3
            Set p_outcome = TestSocketShouldAwaitOperationCompletion
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
''' <remarks>
''' <code>
'''     '++eos' set to 3 in 7.8ms.
'''     '++eoi' set to 1 in 6.2ms.
'''     '++auto' set to 0 in 5.5ms.
'''     '++read_tmo_ms' set to 3000 in 5.6ms.
'''     Serial Poll is 16 in 12.7ms.
'''     Serial Poll is 0 in 3.5ms.
'''     Serial Poll is 16 in 11.4ms.
'''     Serial Poll is 0 in 3.3ms.
''' Test 01 TestSocketShouldConnect passed. Elapsed time: 23.7 ms.
'''     Serial Poll is 16 in 11.2ms.
'''     '++auto' set to 0 in 5.4ms.
'''     '++eos' set to 3 in 6.4ms.
'''     '++eoi' set to 1 in 5.6ms.
'''     '++auto' set to 0 in 5.5ms.
'''     '++read_tmo_ms' set to 3000 in 6.0ms.
'''     Serial Poll is 16 in 26.3ms.
'''     Serial Poll is 0 in 3.4ms.
'''     Serial Poll is 16 in 8.6ms.
'''     Serial Poll is 0 in 3.3ms.
''' Test 02 TestSocketShouldQueryIdentity passed. Elapsed time: 32.3 ms.
'''     Serial Poll is 16 in 14.8ms.
'''     '++auto' set to 0 in 5.6ms.
'''     '++eos' set to 3 in 5.4ms.
'''     '++eoi' set to 1 in 5.6ms.
'''     '++auto' set to 0 in 5.5ms.
'''     '++read_tmo_ms' set to 3000 in 6.7ms.
'''     Serial Poll is 16 in 15.9ms.
'''     Serial Poll is 0 in 3.5ms.
'''     Serial Poll is 96 in 3.8ms.
'''     Serial Poll is 32 in 3.8ms.
''' Test 03 TestSocketShouldAwaitOperationCompletion passed. Elapsed time: 34.7 ms.
'''     Serial Poll is 16 in 14.8ms.
'''     '++auto' set to 0 in 5.5ms.
''' Ran 3 out of 3 tests.
''' Passed: 3; Failed: 0; Inconclusive: 0.
''' </code>
''' </remarks>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 3
    Dim p_testNumber As Integer
    For p_testNumber = 1 To This.TestCount
        Set p_outcome = RunTest(p_testNumber)
        If Not p_outcome Is Nothing Then
            This.RunCount = This.RunCount + 1
            If p_outcome.AssertInconclusive Then
                This.InconclusiveCount = This.InconclusiveCount + 1
            ElseIf p_outcome.AssertSuccessful Then
                This.PassedCount = This.PassedCount + 1
            Else
                This.FailedCount = This.FailedCount + 1
            End If
        End If
        DoEvents
    Next p_testNumber
    AfterAll
    Debug.Print "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
                "; Inconclusive: " & VBA.CStr(This.InconclusiveCount) & "."
End Sub

''' <summary>   Prepares all tests. </summary>
Public Sub BeforeAll()

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed to run all tests.")
    
    This.Name = "SocketSerialPollQueryTests"
    
    This.TestNumber = 0
    This.PreviousTestNumber = 0
    
    This.Address = "192.168.0.252:1234"
    This.PrologixPort = 1234
    This.ReceiveTimeout = 3000
    This.ReadAfterWriteDelay = 1
    This.Termination = VBA.vbLf
    
    ' set to false when testing with serial poll
    This.AssertTalkOnWrite = False
    
    This.IdentityCompany = "KEITHLEY INSTRUMENTS INC."
    
    Set This.ErrTracer = New ErrTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' prime all tests
    
    Set This.DelayStopper = cc_isr_Core_IO.Factory.NewStopwatch
    Set This.TestStopper = cc_isr_Core_IO.Factory.NewStopwatch
        
    Set This.Socket = cc_isr_Winsock.Factory.NewIPv4StreamSocket()
    
    Dim p_details As String
    If Not This.Socket.TryOpenConnection(This.Address, This.ReceiveTimeout, p_details) Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed to run all tests.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Failed priming all tests;" & _
                VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeAllAssert = p_outcome
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
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

    If This.TestNumber = This.PreviousTestNumber Then _
        This.TestNumber = This.PreviousTestNumber + 1

    Dim p_outcome As cc_isr_Test_Fx.Assert

    If This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = IIf(This.Socket.Connected, _
            cc_isr_Test_Fx.Assert.Pass("Ready to prime pre-test #" & VBA.CStr(This.TestNumber) & _
                "; IPV4 Stream Client is connected."), _
            cc_isr_Test_Fx.Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
                ";" & " IPV4 Stream Client should be connected"))
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If

    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    ' Prepare the next test

    Dim p_command As String
    Dim p_sentCount As Integer

    If p_outcome.AssertSuccessful And This.Socket.Port = This.PrologixPort Then
    
        ' prime the GPIB-lan controller
        '
        ' EOS and EOI were set per these recommendations:
        '
        ' https://groups.io/g/HP-Agilent-Keysight-equipment/topic/86224398
        '
        ' set the GPIB termination characters to none - do not append termination characters.
        Set p_outcome = AssertShouldValidateQuery("++eos", "3")
        
        ' Enable EOI assertion with last character
        Set p_outcome = AssertShouldValidateQuery("++eoi", "1")
       
        ' set the read-after-write feature to true.
        Set p_outcome = AssertShouldValidateQuery("++auto", IIf(This.AssertTalkOnWrite, "1", "0"))
    
        If p_outcome.AssertSuccessful Then
        
            Dim p_timeout As Long
            p_timeout = cc_isr_Core_IO.CoreExtensions.ClampLong(This.ReceiveTimeout, 1, 3000)
            Set p_outcome = AssertShouldValidateQuery("++read_tmo_ms", VBA.CStr(p_timeout))
            
        End If
        
        If p_outcome.AssertSuccessful Then
            
            ' disable front panel operation of the currently addressed instrument.
            p_sentCount = This.Socket.SendMessage("++llo" & This.Termination)
            This.DelayStopper.Wait This.ReadAfterWriteDelay
        End If
    
    End If
    
    If p_outcome.AssertSuccessful Then
        
        ' clear execution state before each test.
        ' clear errors if any so as to leave the instrument without errors.
        ' here we add *OPC? to prevent the query unterminated error.
    
        p_sentCount = This.Socket.SendMessage("*CLS;*WAI;*OPC?" & This.Termination)
        This.DelayStopper.Wait This.ReadAfterWriteDelay
        
        If p_outcome.AssertSuccessful And This.Socket.Port = This.PrologixPort Then
        
            Dim p_serialPollOutcome As cc_isr_Test_Fx.Assert
            Set p_serialPollOutcome = AssertSerialPollShouldValidate(16, 16)
            If Not p_serialPollOutcome.AssertSuccessful Then
                Debug.Print p_serialPollOutcome.AssertMessage
            End If
        
        End If
        
    End If
    
    Dim p_reply As String
    Dim p_details As String: p_details = VBA.vbNullString
    If p_outcome.AssertSuccessful Then
        If 0 > TryReceive(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reply, _
            "Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            "; Operation completion query should return the correct reply.")
    
    If p_outcome.AssertSuccessful And This.Socket.Port = This.PrologixPort Then
    
        Set p_serialPollOutcome = AssertSerialPollShouldValidate(0, 16)
        If Not p_serialPollOutcome.AssertSuccessful Then
            Debug.Print p_serialPollOutcome.AssertMessage
        End If
    
    End If
    
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
             Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber))
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                ";" & VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeEachAssert = p_outcome

    On Error GoTo 0
    
    This.TestStopper.Restart
    
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases test elements after each tests is run. </summary>
Public Sub AfterEach()
    
    Const p_procedureName As String = "AfterEach"
    
    ' Trap errors to the error handler.
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")

    If Not This.BeforeEachAssert.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to cleanup test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeEachAssert.AssertMessage)

    ' cleanup after each test.
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_command As String
        Dim p_sentCount As Integer
        Dim p_reply As String
    
        ' clear errors if any so as to leave the instrument without errors.
        p_command = "*CLS;*WAI;*OPC?"
        p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
        This.DelayStopper.Wait This.ReadAfterWriteDelay
        
        If p_outcome.AssertSuccessful And This.Socket.Port = This.PrologixPort Then
        
            Dim p_serialPollOutcome As cc_isr_Test_Fx.Assert
            Set p_serialPollOutcome = AssertSerialPollShouldValidate(16, 16)
            If Not p_serialPollOutcome.AssertSuccessful Then
                Debug.Print p_serialPollOutcome.AssertMessage
            End If
        
        End If
        
        Dim p_details As String: p_details = VBA.vbNullString
        If 0 > TryReceive(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        This.DelayStopper.Wait This.ReadAfterWriteDelay
        
    End If
        
    ' Restore GPIB Lan Controller state
    If p_outcome.AssertSuccessful And This.Socket.Port = This.PrologixPort Then
    
        ' set the read-after-write feature to false.
        Set p_outcome = AssertShouldValidateQuery("++auto", "0")
    
        ' restore front panel operation of the currently addressed instrument.
        
        p_sentCount = This.Socket.SendMessage("++loc" & This.Termination)
        This.DelayStopper.Wait This.ReadAfterWriteDelay

    End If
        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:
    
    ' record the previous test number
    This.PreviousTestNumber = This.TestNumber

    ' release the 'Before Each' cc_isr_Test_Fx.Assert.
    Set This.BeforeEachAssert = Nothing

    If p_outcome.AssertSuccessful Then
    
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Errors reported cleaning up test #" & VBA.CStr(This.TestNumber) & _
                ";" & VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    
    End If

    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
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
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = cc_isr_Test_Fx.Assert.Pass("All tests cleaned up.")
    
    ' cleanup after all tests.
    
    ' disconnect if connected
    Dim p_details As String: p_details = VBA.vbNullString
    If Not This.Socket Is Nothing Then
        If Not This.Socket.TryCloseConnection(p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
        
    Set This.Socket = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before All' cc_isr_Test_Fx.Assert.
    Set This.BeforeAllAssert = Nothing

    ' report any leftover errors.
    Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Errors reported cleaning up all tests;" & _
            VBA.vbCrLf & p_outcome.AssertMessage)
    End If
    
    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

Public Function TryReceive(ByRef a_reply As String, ByRef a_details As String) As Integer

    Dim p_command As String
    If Not This.AssertTalkOnWrite Then
        p_command = "++read eoi"
        Dim p_sentCount As Integer
        p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
        This.DelayStopper.Wait This.ReadAfterWriteDelay
    End If
    
    TryReceive = This.Socket.TryReceive(a_reply, a_details)

End Function

Public Function TryQuery(ByVal a_command As String, ByRef a_reply As String, ByRef a_details As String) As Integer
    
    ' send the command
    Dim p_sentCount As Integer
    p_sentCount = This.Socket.SendMessage(a_command & This.Termination)
    This.DelayStopper.Wait This.ReadAfterWriteDelay
    
    If This.Socket.Port = This.PrologixPort Then
    
        Dim p_serialPollOutcome As cc_isr_Test_Fx.Assert
        Set p_serialPollOutcome = AssertSerialPollShouldValidate(16, 16)
        If Not p_serialPollOutcome.AssertSuccessful Then
            Debug.Print p_serialPollOutcome.AssertMessage
        End If
    
    End If
    
    TryQuery = TryReceive(a_reply, a_details)

End Function

''' summary>   Asserts that the status byte bits value are correct. </summary>
''' <param name="a_bitsStatus"/>   [Integer] The expected status of the specified status bits. </param>
''' <param name="a_statusBits"/>   [Integer] The expected status bits. </param>
Private Function AssertSerialPollShouldValidate(ByVal a_bitsStatus As Integer, ByVal a_statusBits As Integer) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    If This.AssertTalkOnWrite Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass()
    Else
        Dim p_details As String
        Dim p_polled As Boolean
        Dim p_elapsed As Double
        Dim p_statusByte As Integer
        Dim p_stopper As cc_isr_Core_IO.Stopwatch
        Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
        p_stopper.Restart
        p_polled = AwaitStatusBits(a_bitsStatus, a_statusBits, 3000, p_statusByte, p_details)
        p_elapsed = p_stopper.ElapsedMilliseconds
        If p_statusByte < 0 Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        ElseIf p_polled Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Pass()
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail("    Status byte '" & _
                VBA.CStr(p_statusByte) & "' bits '" & VBA.CStr(a_statusBits) & _
                "' not matching the expected bits '" & VBA.CStr(a_bitsStatus) & "' value.")
        End If
        Debug.Print "    Serial Poll is " & VBA.CStr(p_statusByte) & _
            " in " & Format(p_elapsed, "0.0") & "ms."
    End If
    Set AssertSerialPollShouldValidate = p_outcome

End Function

''' summary>   Reads the status byte of the current GPIB instrument. </summary>
''' <param name="a_details">       [Out, String] details the failure reason. </param>
''' <returns>   [Integer] The status byte or RECEIVE_ERROR (-1) if failed receiving a reading or
''' failed parsing the reading to an integer.
''' </returns>
Public Function SerialPoll(ByRef a_details As String) As Integer
    
    Dim p_command As String
    p_command = "++spoll"
    
    Dim p_sentCount As Integer
    p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
    This.DelayStopper.Wait This.ReadAfterWriteDelay
    
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
        a_statusByte = SerialPoll(a_details)
        p_gotIt = (a_statusByte >= 0) And (a_bitsStatus = (a_statusBits And a_statusByte))
    Loop Until p_gotIt Or (a_statusByte < 0) Or (p_stopper.ElapsedMilliseconds > a_timeout)
    
    AwaitStatusBits = p_gotIt
    
End Function

Public Function IsMessageAvailable(ByVal a_MAV As Integer, ByRef a_statusByte As Integer, ByRef a_details As String) As Boolean
    
    a_statusByte = SerialPoll(a_details)
    If a_statusByte > 0 Then
        IsMessageAvailable = (a_MAV - (a_MAV And a_statusByte))
    Else
        IsMessageAvailable = False
    End If
End Function

Public Function AwaitMessageAvailable(ByVal a_MAV As Integer, ByVal timeout As Integer, _
    ByRef a_statusByte As Integer, ByRef a_details As String) As Boolean

    Dim p_stopper As cc_isr_Core_IO.Stopwatch
    p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
    p_stopper.Restart
    Dim p_messageAvailable As Boolean
    p_messageAvailable = IsMessageAvailable(a_MAV, a_statusByte, a_details)
    While Not p_messageAvailable And (p_stopper.ElapsedMilliseconds < timeout)
        DoEvents
        p_messageAvailable = IsMessageAvailable(a_MAV, a_statusByte, a_details)
    Wend
    AwaitMessageAvailable = p_messageAvailable
    
End Function

''' <summary>   Asserts a valid serial poll. </summary>
Private Function AssertSerialPollShouldValidate_(ByVal a_value As Integer, ByVal a_bitValue As Integer, _
    ByRef a_statusByte As Integer) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_receiveCount As Integer
    Dim p_reply As String
    Dim p_details As String: p_details = VBA.vbNullString
    
    p_command = "++spoll"
    
    ' send the command
    p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
    This.DelayStopper.Wait This.ReadAfterWriteDelay

    ' the receive count is negative if error
    p_receiveCount = This.Socket.TryReceive(p_reply, p_details)
    If 0 > p_receiveCount Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass()
    End If
    This.DelayStopper.Wait This.ReadAfterWriteDelay
    
    If p_outcome.AssertSuccessful Then
    
        If Not cc_isr_core.StringExtensions.TryParseInteger(p_reply, a_statusByte, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_value, a_bitValue And a_statusByte, _
                "    Status byte '" & VBA.CStr(a_statusByte) & "' bits not matching the expected value.")
        End If
    End If
            
    Set AssertSerialPollShouldValidate_ = p_outcome
End Function

Private Function AssertShouldValidateQuery(ByVal a_command As String, ByVal a_value As String) As cc_isr_Test_Fx.Assert
    Dim p_elapsed As Double
    Dim p_stopper As cc_isr_Core_IO.Stopwatch
    Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
    p_stopper.Restart
    Set AssertShouldValidateQuery = AssertShouldValidateQuery_(a_command, a_value)
    p_elapsed = p_stopper.ElapsedMilliseconds
    Debug.Print "    '" & a_command & "' set to " & VBA.CStr(a_value) & _
        " in " & Format(p_elapsed, "0.0") & "ms."
End Function

Private Function AssertShouldValidateQuery_(ByVal a_command As String, ByVal a_value As String) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_receiveCount As Integer
    Dim p_reply As String
    Dim p_details As String: p_details = VBA.vbNullString
    
    ' set auto read after write
    p_command = a_command & " " & a_value

    ' send the command
    p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
    This.DelayStopper.Wait This.ReadAfterWriteDelay

    ' validate reading
    
    ' set auto query command
    p_command = a_command
    
    p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
    This.DelayStopper.Wait This.ReadAfterWriteDelay
    
    ' the receive count is negative if error
    p_receiveCount = This.Socket.TryReceive(p_reply, p_details)
    If 0 > p_receiveCount Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass()
    End If
    This.DelayStopper.Wait This.ReadAfterWriteDelay
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_value, p_reply, _
            " Command '" & a_command & "' value does not match its expected.")
            
    Set AssertShouldValidateQuery_ = p_outcome

End Function

''' <summary>   Unit test. Asserts that the stream socket should query a device identity. </summary>
''' <remarks>
''' <code>
''' With 5ms delay
''' '++eos' set to 3 in 15.6ms.
''' '++eoi' set to 1 in 15.5ms.
''' '++auto' set to 0 in 15.5ms.
''' '++read_tmo_ms' set to 3000 in 15.6ms.
''' Serial Poll is 16 in 10.8ms.
''' Serial Poll is 0 in 8.5ms.
''' Serial Poll is 16 in 10.7ms.
''' Serial Poll is 0 in 5.3ms.
''' TestSocketShouldConnect passed.
''' Serial Poll is 16 in 10.7ms.
'''
''' '++auto' set to 0 in 15.5ms.
''' '++eos' set to 3 in 19.5ms.
''' '++eoi' set to 1 in 17.3ms.
''' '++auto' set to 0 in 15.6ms.
''' '++read_tmo_ms' set to 3000 in 15.6ms.
''' Serial Poll is 16 in 21.4ms.
''' Serial Poll is 0 in 5.5ms.
''' Serial Poll is 16 in 10.7ms.
''' Serial Poll is 0 in 5.5ms.
''' TestSocketShouldConnect passed.
''' Serial Poll is 16 in 10.7ms.
''' '++auto' set to 0 in 15.5ms.
'''
''' With 1ms delay
''' '++eos' set to 3 in 6.2ms.
''' '++eoi' set to 1 in 3.5ms.
''' '++auto' set to 0 in 3.7ms.
''' '++read_tmo_ms' set to 3000 in 4.8ms.
''' Serial Poll is 16 in 15.0ms.
''' Serial Poll is 0 in 3.5ms.
''' Serial Poll is 16 in 13.5ms.
''' Serial Poll is 0 in 4.7ms.
''' TestSocketShouldConnect passed. in 26.4 ms.
''' Serial Poll is 16 in 13.2ms.
''' '++auto' set to 0 in 3.3ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestSocketShouldConnect() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestSocketShouldConnect"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_reply As String
    
    If p_outcome.AssertSuccessful Then
            
        ' check if connected and clear errors.
        p_command = "*CLS;*WAI;*OPC?"
        p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
        This.DelayStopper.Wait This.ReadAfterWriteDelay
        
    End If
    
    If p_outcome.AssertSuccessful And This.Socket.Port = This.PrologixPort Then
    
        Dim p_serialPollOutcome As cc_isr_Test_Fx.Assert
        Set p_serialPollOutcome = AssertSerialPollShouldValidate(16, 16)
        If Not p_serialPollOutcome.AssertSuccessful Then
            Debug.Print p_serialPollOutcome.AssertMessage
        End If
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_details As String: p_details = VBA.vbNullString
        If 0 > TryReceive(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        This.DelayStopper.Wait This.ReadAfterWriteDelay
    
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertSerialPollShouldValidate(0, 16)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestSocketShouldConnect = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts that the stream socket should query a device identity. </summary>
''' <remarks>
''' <code>
''' '++eos' set to 3 in 15.9ms.
''' '++eoi' set to 1 in 15.5ms.
''' '++auto' set to 0 in 15.5ms.
''' '++read_tmo_ms' set to 3000 in 15.6ms.
''' Serial Poll is 16 in 16.0ms.
''' Serial Poll is 0 in 5.4ms.
''' Serial Poll is 16 in 10.7ms.
''' Serial Poll is 0 in 5.4ms.
''' TestSocketShouldQueryIdentity passed.
''' Serial Poll is 16 in 22.3ms.
''' '++auto' set to 0 in 15.4ms.
'''
''' With 1 ms read after write delay.
'''
''' Serial Poll is 16 in 12.5ms.
''' '++auto' set to 0 in 3.5ms.
''' '++eos' set to 3 in 6.9ms.
''' '++eoi' set to 1 in 3.7ms.
''' '++auto' set to 0 in 3.5ms.
''' '++read_tmo_ms' set to 3000 in 4.1ms.
''' Serial Poll is 16 in 21.8ms.
''' Serial Poll is 0 in 3.5ms.
''' Serial Poll is 16 in 9.9ms.
''' Serial Poll is 0 in 3.4ms.
''' TestSocketShouldQueryIdentity passed.
''' Serial Poll is 16 in 12.9ms.
''' '++auto' set to 0 in 3.3ms.
'''
''' '++eos' set to 3 in 5.9ms.
''' '++eoi' set to 1 in 3.8ms.
''' '++auto' set to 0 in 3.9ms.
''' '++read_tmo_ms' set to 3000 in 4.3ms.
''' Serial Poll is 16 in 14.6ms.
''' Serial Poll is 0 in 3.5ms.
''' Serial Poll is 16 in 7.4ms.
''' Serial Poll is 0 in 3.3ms.
''' TestSocketShouldQueryIdentity passed. in 37.3 ms.
''' Serial Poll is 16 in 14.2ms.
''' '++auto' set to 0 in 3.7ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestSocketShouldQueryIdentity() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestSocketShouldQueryIdentity"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String: p_command = "*IDN?"
    Dim p_sentCount As Integer
    Dim p_identity As String
    Dim p_readCount As Integer
    Dim p_reply As String
    
    If p_outcome.AssertSuccessful Then
            
        ' send the command
        p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
        This.DelayStopper.Wait This.ReadAfterWriteDelay
    
    End If

    If p_outcome.AssertSuccessful And This.Socket.Port = This.PrologixPort Then
    
        Dim p_serialPollOutcome As cc_isr_Test_Fx.Assert
        Set p_serialPollOutcome = AssertSerialPollShouldValidate(16, 16)
        If Not p_serialPollOutcome.AssertSuccessful Then
            Debug.Print p_serialPollOutcome.AssertMessage
        End If
    
    End If
    
    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then
        
        If 0 > TryReceive(p_identity, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        This.DelayStopper.Wait This.ReadAfterWriteDelay
    
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
            1 = VBA.InStr(1, p_identity, This.IdentityCompany, VBA.VbCompareMethod.vbTextCompare), _
            "Identity '" & p_identity & " should start with '" & This.IdentityCompany & "'.")

    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertSerialPollShouldValidate(0, 16)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestSocketShouldQueryIdentity = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

''' <summary>   Unit test. Asserts that the stream socket should await operation completion. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' '++eos' set to 3 in 9.5ms.
''' '++eoi' set to 1 in 3.5ms.
''' '++auto' set to 0 in 3.5ms.
''' '++read_tmo_ms' set to 3000 in 4.4ms.
''' Serial Poll is 16 in 17.5ms.
''' Serial Poll is 0 in 3.6ms.
''' Serial Poll is 96 in 5.1ms.
''' Serial Poll is 32 in 4.0ms.
''' TestSocketShouldAwaitOperationCompletion passed. in 49.0 ms.
''' Serial Poll is 16 in 13.9ms.
''' '++auto' set to 0 in 3.8ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestSocketShouldAwaitOperationCompletion() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestSocketShouldAwaitOperationCompletion"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String
    Dim p_sentCount As Integer
    
    If p_outcome.AssertSuccessful Then
            
        ' clear known state, enable OPC Standard Event and Service Request on the standard event bit.
        p_command = "*CLS;*ESE 1;*SRE 32"
        p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
        This.DelayStopper.Wait This.ReadAfterWriteDelay
        
        ' syncrhronize.
        p_command = "*OPC?"
        p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
        This.DelayStopper.Wait This.ReadAfterWriteDelay
        
    End If
    
    Dim p_reply As String
    Dim p_details As String: p_details = VBA.vbNullString
    If p_outcome.AssertSuccessful Then
        If 0 > TryReceive(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reply, _
            "Operation completion query should return the correct reply.")
        
    If p_outcome.AssertSuccessful Then
        
        p_command = "*OPC"
        p_sentCount = This.Socket.SendMessage(p_command & This.Termination)
        This.DelayStopper.Wait This.ReadAfterWriteDelay
        
    End If
        
    ' wait for the operation completion bit.
    Dim p_stadnardEventBit As Integer
    p_stadnardEventBit = 32
    
    Dim p_requestingServiceBit As Integer
    p_requestingServiceBit = 64
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertSerialPollShouldValidate(64, 64)

    ' wait for the operation completion bit to turn off.
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = AssertSerialPollShouldValidate(0, 64)

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestSocketShouldAwaitOperationCompletion = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function



