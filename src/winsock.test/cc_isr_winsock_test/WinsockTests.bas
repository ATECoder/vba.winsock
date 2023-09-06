Attribute VB_Name = "WinsockTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Winsock tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   This class properties. </summary>
Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
    ErrTracer As IErrTracer
End Type

Private This As this_

''' <summary>   Runs the specified test. </summary>
Public Sub RunTest(ByVal a_testNumber As Integer)
    BeforeEach
    Select Case a_testNumber
        Case 1
            TestInitializeAndDispose
        Case 2
            TestGettingLastErrorDescription
        Case Else
    End Select
    AfterEach
End Sub

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    BeforeAll
    Dim p_testNumber As Integer
    For p_testNumber = 1 To 2
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

    This.Name = "WinsockTests"
    
    This.TestNumber = 0
    
    Set This.ErrTracer = New ErrTracer
    
    Set This.BeforeAllAssert = Assert.Pass("initialize the overall assert.")
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' report any leftover archived errors.
    If cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildArchivedErrorsMessage()
        Set This.BeforeAllAssert = Assert.Inconclusive("Failed preparing all tests: " & _
            p_leftoverErrorMessage)
        This.ErrTracer.TraceError p_leftoverErrorMessage
    
    End If

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


    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
        
        Set This.BeforeEachAssert = Assert.Pass("initialize the pre-test assert.")
    
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
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' report any leftover archived errors.
    If cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildArchivedErrorsMessage()
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
    
    ' enqueue the error or append its source to the last error.
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

    Set This.BeforeEachAssert = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' report any leftover archived errors.
    If cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildArchivedErrorsMessage()
        This.ErrTracer.TraceError "Error(s) were stacked unwinding test #" & _
            VBA.CStr(This.TestNumber) & ": " & p_leftoverErrorMessage
    
    End If

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
    
    Set This.BeforeAllAssert = Nothing

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' report any leftover archived errors.
    If cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildArchivedErrorsMessage()
        This.ErrTracer.TraceError "Errors were stacked unwinding all tests: " & p_leftoverErrorMessage
    
    End If

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

''' <summary>   Unit test. Asserts instantiating and disposing of the Winsock framework. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInitializeAndDispose() As Assert

    Const p_procedureName As String = "TestInitializeAndDispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    ' this is required to initialize Winsock.  It will only ran once.
    Winsock.Initialize
    
    Set p_outcome = Assert.IsTrue(Winsock.Initiated, "Winsock should be initiated when a socket is created")
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsFalse(Winsock.Disposed, "Winsock should not be disposed when a socket is created")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(Winsock.SocketCount, 0, _
            "Winsock socket count should be 0 as no sockets are registered but is " & Str$(Winsock.SocketCount))
    End If

    ' test disposing of Winsock.
    Winsock.Dispose
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsFalse(Winsock.Initiated, _
            "Winsock should no longer be initiated after the last socket was set to nothing")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(Winsock.Disposed, _
            "Winsock should be disposed after the last socket was set to nothing")
    End If
    
    Winsock.Dispose
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestInitializeAndDispose")
    
    Set TestInitializeAndDispose = p_outcome
    
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

''' <summary>   Unit test. Asserts getting the error description from the Windows API. </summary>
''' <returns>   An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestGettingLastErrorDescription() As Assert

    Const p_procedureName As String = "TestGettingLastErrorDescription"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_errorNumber As Long: p_errorNumber = 5
    Dim p_expected As String: p_expected = "Access is denied."
    Dim p_actual As String: p_actual = Winsock.LastErrorDescription(p_errorNumber)
    
    Set p_outcome = Assert.AreEqual(p_expected, p_actual, _
            "Winsock should get the correct error description for error number " & CStr(p_errorNumber))
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestGettingLastErrorDescription")
    
    Set TestGettingLastErrorDescription = p_outcome
    
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













