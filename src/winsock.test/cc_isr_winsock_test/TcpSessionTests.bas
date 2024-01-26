Attribute VB_Name = "TcpSessionTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Tcp Session query identity Tests. </summary>
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
    Session As TcpSession
    ReceiveTimeout As Long
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
            Set p_outcome = TestShouldConnect
        Case 2
            Set p_outcome = TestShouldQueryIdentity
        Case 3
            Set p_outcome = TestShouldAwaitOperationCompletion
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 3
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
''' <remarks>
''' <code>
''' Test 01 TestShouldConnect passed. Elapsed time: 13.8 ms.
''' Test 02 TestShouldQueryIdentity passed. Elapsed time: 20.9 ms.
''' Test 03 TestShouldAwaitOperationCompletion passed. Elapsed time: 34.4 ms.
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
    
    This.Name = "TcpSessionQueryTests"
    
    This.TestNumber = 0
    This.PreviousTestNumber = 0
    
    This.Address = "192.168.0.252:1234"
    This.ReceiveTimeout = 3000
    
    ' set to false when testing with serial poll
    This.AssertTalkOnWrite = False
    
    This.IdentityCompany = "KEITHLEY INSTRUMENTS INC."
    
    Set This.ErrTracer = New ErrTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' prime all tests
    
    Set This.DelayStopper = cc_isr_Core_IO.Factory.NewStopwatch
    Set This.TestStopper = cc_isr_Core_IO.Factory.NewStopwatch
        
    Set This.Session = New TcpSession
    This.Session.Initialize cc_isr_Winsock.Factory.NewIPv4StreamSocket()
    This.Session.GpibLanControllerPort = 1234
    This.Session.Termination = VBA.vbLf
    This.Session.ReadAfterWriteDelay = 1
   
    Dim p_details As String
    If Not This.Session.Socket.TryOpenConnection(This.Address, This.ReceiveTimeout, p_details) Then
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
        Set p_outcome = IIf(This.Session.Socket.Connected, _
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
   
    If p_outcome.AssertSuccessful Then
        
        ' clear execution state before each test.
        ' clear errors if any so as to leave the instrument without errors.
        ' here we add *OPC? to prevent the query unterminated error.
    
        This.Session.SendMessage ("*CLS;*WAI;*OPC?")
       
    End If
    
    Dim p_reply As String
    Dim p_details As String: p_details = VBA.vbNullString
    If p_outcome.AssertSuccessful Then
        If 0 > This.Session.TryReceive(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reply, _
            "Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            "; Operation completion query should return the correct reply.")
    
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
        Dim p_reply As String
    
        ' clear errors if any so as to leave the instrument without errors.
        p_command = "*CLS;*WAI;*OPC?"
        This.Session.SendMessage (p_command)
        
        Dim p_details As String: p_details = VBA.vbNullString
        If 0 > This.Session.TryReceive(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
        
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
    If Not This.Session Is Nothing Then
        If Not This.Session.Socket Is Nothing Then
            If Not This.Session.Socket.TryCloseConnection(p_details) Then
                Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
            End If
        End If
    End If
        
    Set This.Session = Nothing

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

Private Function AssertShouldValidateQuery(ByVal a_command As String, ByVal a_value As String) As cc_isr_Test_Fx.Assert
    Dim p_elapsed As Double
    Dim p_stopper As cc_isr_Core_IO.Stopwatch
    Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch()
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_result As String
    Dim p_details As String
    Dim p_setCommand As String
    p_setCommand = a_command & " " & a_value
    p_stopper.Restart
    This.Session.SendMessage (a_command & " " & a_value)
    
    If This.Session.TryGetValue(a_command, a_value, p_result, p_details) Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass()
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
    End If
    p_elapsed = p_stopper.ElapsedMilliseconds
    Set AssertShouldValidateQuery = p_outcome
    Debug.Print "    '" & p_setCommand & "' value set to " & p_result & _
        " in " & Format(p_elapsed, "0.0") & "ms."
End Function

''' <summary>   Unit test. Asserts that the session should query a device identity. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay.
''' TestShouldConnect passed. in 13.0 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldConnect() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldConnect"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_reply As String
    
    If p_outcome.AssertSuccessful Then
            
        ' check if connected and clear errors.
        p_command = "*CLS;*WAI;*OPC?"
        p_sentCount = This.Session.SendMessage(p_command)
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_details As String: p_details = VBA.vbNullString
        If 0 > This.Session.TryReceive(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldConnect = p_outcome
    
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

''' <summary>   Unit test. Asserts that the session should query a device identity. </summary>
''' <remarks>
''' <code>
''' With 1ms read after write delay
''' TestShouldQueryIdentity passed. in 20.3 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldQueryIdentity() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldQueryIdentity"
    
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
        p_sentCount = This.Session.SendMessage(p_command)
    
    End If

    Dim p_details As String: p_details = VBA.vbNullString
    
    If p_outcome.AssertSuccessful Then
        
        If 0 > This.Session.TryReceive(p_identity, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    
    End If
    
    If p_outcome.AssertSuccessful Then
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
            1 = VBA.InStr(1, p_identity, This.IdentityCompany, VBA.VbCompareMethod.vbTextCompare), _
            "Identity '" & p_identity & " should start with '" & This.IdentityCompany & "'.")

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldQueryIdentity = p_outcome
    
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

''' <summary>   Unit test. Asserts that the session should await operation completion. </summary>
''' <remarks>
''' <code>
''' With 1ms read sfter write delay.
''' TestShouldAwaitOperationCompletion passed. in 45.2 ms.
''' </code>
''' </remarks>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestShouldAwaitOperationCompletion() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestShouldAwaitOperationCompletion"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String
    Dim p_sentCount As Integer
    
    If p_outcome.AssertSuccessful Then
            
        ' clear known state, enable OPC Standard Event and Service Request on the standard event bit.
        p_command = "*CLS;*ESE 1;*SRE 32"
        p_sentCount = This.Session.SendMessage(p_command)
        
        ' syncrhronize.
        p_command = "*OPC?"
        p_sentCount = This.Session.SendMessage(p_command)
        
    End If
    
    Dim p_reply As String
    Dim p_details As String: p_details = VBA.vbNullString
    If p_outcome.AssertSuccessful Then
        If 0 > This.Session.TryReceive(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reply, _
            "Operation completion query should return the correct reply.")
        
    If p_outcome.AssertSuccessful Then
        
        p_command = "*OPC"
        p_sentCount = This.Session.SendMessage(p_command)
        
        ' read status byte
        p_command = "*STB?"
        p_sentCount = This.Session.SendMessage(p_command)
        
    End If
        
    If p_outcome.AssertSuccessful Then
        If 0 > This.Session.TryReceive(p_reply, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    Dim p_statusByte As Integer
    If p_outcome.AssertSuccessful Then
        If Not cc_isr_core.StringExtensions.TryParseInteger(p_reply, p_statusByte, p_details) Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Fail(p_details)
        End If
    End If
    
    ' wait for the operation completion bit.
    Dim p_stadnardEventBit As Integer
    p_stadnardEventBit = 32
    
    Dim p_requestingServiceBit As Integer
    p_requestingServiceBit = 64
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_requestingServiceBit, _
            p_requestingServiceBit And p_statusByte, _
            "Status byte '" & VBA.CStr(p_statusByte) & _
            "' requesting service bit 6 '" & VBA.CStr(p_requestingServiceBit) & "' should be set.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestShouldAwaitOperationCompletion = p_outcome
    
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







