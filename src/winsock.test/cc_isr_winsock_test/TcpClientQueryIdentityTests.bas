Attribute VB_Name = "TcpClientQueryIdentityTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Client query identity Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   This class properties. </summary>
Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    Host As String
    Port As Long
    PrologixPort As Long
    Client As TcpClient
    ReceiveTimeout As Long
    ReadAfterWriteDelay As Integer
    AssertTalkOnWrite As Boolean
    DelayStopper As cc_isr_Core_IO.Stopwatch
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
    BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestTcpClentShouldQueryIdentity
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
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 1
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

    This.Name = "TcpClientQueryIdentityTests"

    This.TestNumber = 0
    This.Host = "192.168.0.252"
    This.Port = 1234
    This.PrologixPort = 1234
    This.ReceiveTimeout = 3000
    
    This.ReceiveTimeout = 3000
    This.ReadAfterWriteDelay = 1
    
    ' set to false when testing with serial poll
    This.AssertTalkOnWrite = False
    
    This.IdentityCompany = "KEITHLEY INSTRUMENTS INC."

    Set This.DelayStopper = cc_isr_Core_IO.Factory.NewStopwatch
        
    Set This.ErrTracer = New ErrTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    ' prime all tests
    
    Set This.Client = cc_isr_Winsock.Factory.NewTcpClient()
    
    This.Client.OpenConnection This.Host, This.Port, This.ReceiveTimeout
   
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

    This.TestNumber = This.TestNumber + 1

    Dim p_outcome As cc_isr_Test_Fx.Assert

    If This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = IIf(This.Client.Connected, _
            cc_isr_Test_Fx.Assert.Pass("Ready to prime pre-test #" & VBA.CStr(This.TestNumber) & _
                "; TCP Client is connected."), _
            cc_isr_Test_Fx.Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
                ";" & " TCP Client should be connected"))
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    ' Prepare the next test
    
    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_reply As String

    If p_outcome.AssertSuccessful Then
    
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
            p_sentCount = This.Client.SendMessage(p_command & VBA.vbLf)
            This.DelayStopper.Wait 5
    
            ' disables front panel operation of the currently addressed instrument.
            
            p_sentCount = This.Client.SendMessage("++llo" & VBA.vbLf)
            This.DelayStopper.Wait 5
    
        End If
    
        ' clear execution state before each test.
        ' clear errors if any so as to leave the instrument without errors.
        ' here we add *OPC? to prevent the query unterminated error.
    
        p_sentCount = This.Client.SendMessage("*CLS;*WAI;*OPC?" & VBA.vbLf)
        This.DelayStopper.Wait 5
        p_reply = This.Client.ReceiveRaw
        This.DelayStopper.Wait 5
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("1", p_reply, _
                "Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
                "; Operation completion query should send the correct reply.")
            
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

    ' cleanup after each test.
    If This.BeforeEachAssert.AssertSuccessful Then
    
        Dim p_command As String
        Dim p_sentCount As Integer
        Dim p_reply As String
    
        ' clear errors if any so as to leave the instrument without errors.
        p_sentCount = This.Client.SendMessage("*CLS;*WAI;*OPC?" & VBA.vbLf)
        This.DelayStopper.Wait 5
        p_reply = This.Client.ReceiveRaw
        This.DelayStopper.Wait 5

        ' Restore Prologix device.
        If This.BeforeEachAssert.AssertSuccessful And This.Port = This.PrologixPort Then
        
            p_command = "++auto 0"

            ' send the command, which may cause Query Unterminated because we are setting the device to talk
            ' where there is nothing to talk.
            p_sentCount = This.Client.SendMessage(p_command & VBA.vbLf)
            This.DelayStopper.Wait 5

            ' restore front panel operation of the currently addressed instrument.
            
            p_sentCount = This.Client.SendMessage("++loc" & VBA.vbLf)
            This.DelayStopper.Wait 5

        End If

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before Each' cc_isr_Test_Fx.Assert.
    Set This.BeforeEachAssert = Nothing

    ' report any leftover errors.
    Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Errors reported cleaning up test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & p_outcome.AssertMessage)
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
    If Not This.Client Is Nothing Then _
        This.Client.CloseConnection

    Set This.Client = Nothing

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

''' <summary>   Unit test. Asserts that the TCP Client should query a device identity. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function TestTcpClentShouldQueryIdentity() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTcpClentShouldQueryIdentity"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String: p_command = "*IDN?"
    Dim p_sentCount As Integer
    Dim p_identity As String
    
    If p_outcome.AssertSuccessful Then
            
        ' send the command
        p_sentCount = This.Client.SendMessage(p_command & VBA.vbLf)
        This.DelayStopper.Wait 5
            
        p_identity = This.Client.ReceiveRaw()

    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestTcpClentShouldQueryIdentity")
    
    Set TestTcpClentShouldQueryIdentity = p_outcome
    
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

