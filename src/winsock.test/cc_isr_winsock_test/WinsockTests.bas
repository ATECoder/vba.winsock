Attribute VB_Name = "WinsockTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Winsock tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
    ErrTracer As IErrTracer
End Type

Private This As this_

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

Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

Public Sub RunAllTests()
    BeforeAll
    Dim p_testNumber As Integer
    For p_testNumber = 1 To 2
        RunTest p_testNumber
        DoEvents
    Next p_testNumber
    AfterAll
End Sub

Public Sub BeforeAll()

    This.TestNumber = 0
    Set This.BeforeAllAssert = Assert.Pass("initialize the overall assert.")
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    Set This.ErrTracer = New ErrTracer
    
End Sub

Public Sub BeforeEach()

    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
        
        Set This.BeforeEachAssert = Assert.Pass("initialize the pre-test assert.")
    
    Else
    
        Set This.BeforeEachAssert = Assert.Inconclusive(This.BeforeAllAssert.AssertMessage)
    
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    If This.BeforeEachAssert.AssertSuccessful Then
    
        Set This.BeforeEachAssert = Assert.AreEqual(0, Err.Number, _
            "Error Number should be 0.")
            
    End If
    
    This.TestNumber = This.TestNumber + 1
                        
End Sub

Public Sub AfterEach()
    
    Set This.BeforeEachAssert = Nothing

End Sub

Public Sub AfterAll()
    
    Set This.BeforeAllAssert = Nothing

End Sub

''' <summary>   Unit test. Asserts instantiating and disposing of the Winsock framework. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInitializeAndDispose() As Assert

    Dim p_outcome As Assert

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
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestInitializeAndDispose")
    
    Set TestInitializeAndDispose = p_outcome
    
End Function

''' <summary>   Unit test. Asserts getting the error description from the Windows API. </summary>
''' <returns>   An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestGettingLastErrorDescription() As Assert

    Dim p_outcome As Assert
    Dim p_errorNumber As Long: p_errorNumber = 5
    Dim p_expected As String: p_expected = "Access is denied."
    Dim p_actual As String: p_actual = Winsock.LastErrorDescription(p_errorNumber)
    
    Set p_outcome = Assert.AreEqual(p_expected, p_actual, _
            "Winsock should get the correct error description for error number " & CStr(p_errorNumber))
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestGettingLastErrorDescription")
    
    Set TestGettingLastErrorDescription = p_outcome
    
End Function













