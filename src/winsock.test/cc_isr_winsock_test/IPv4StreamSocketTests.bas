Attribute VB_Name = "IPv4StreamSocketTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   IPV4 Stream socket tests. </summary>
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
            TestCreateSocket
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
    For p_testNumber = 1 To 1
        RunTest p_testNumber
        DoEvents
    Next p_testNumber
    AfterAll
End Sub

Public Sub BeforeAll()

    This.TestNumber = 0
    Set This.BeforeAllAssert = Assert.IsTrue(True, "initialize the overall assert.")
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    Set This.ErrTracer = New ErrTracer
    
End Sub

Public Sub BeforeEach()

    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
        
        Set This.BeforeEachAssert = Assert.IsTrue(True, "initialize the pre-test assert.")
    
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


''' <summary>   Unit test. Asserts creating a socket. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCreateSocket() As Assert

    Dim p_outcome As Assert
    
    Dim p_socket As IPv4StreamSocket
    Set p_socket = cc_isr_Winsock.Factory.NewIPv4StreamSocket
    
    ' check if socket has a valid id
    Set p_outcome = Assert.IsTrue(p_socket.SocketId <> wsock32.ws32_INVALID_SOCKET, _
        "Failed creating socket; socket id " & Str$(p_socket.SocketId) & _
        " must not equal to wsock32.INVALID_SOCKET=" & wsock32.ws32_INVALID_SOCKET)
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(Winsock.Initiated, _
            "Winsock should be initiated when a socket is created")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsFalse(Winsock.Disposed, "Winsock should not be disposed when a socket is created")
    End If
    
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(Winsock.SocketCount, 1, _
            "Winsock socket count should be 1 after registering a single socket but is " & Str$(Winsock.SocketCount))
    End If
    
    ' test terminating the socket, which should dispose of the Winsock class.
    Set p_socket = Nothing
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.AreEqual(Winsock.SocketCount, 0, _
            "Winsock socket count should be 0 after nulling single socket but is " & Str$(Winsock.SocketCount))
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsFalse(Winsock.Initiated, _
            "Winsock should no longer be initiated after the last socket was set to nothing")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.IsTrue(Winsock.Disposed, _
            "Winsock should be disposed after the last socket was set to nothing")
    End If
       
    Set p_socket = Nothing
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestCreateSocket")
    
    Set TestCreateSocket = p_outcome
    
End Function




