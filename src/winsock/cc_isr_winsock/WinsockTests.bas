Attribute VB_Name = "WinsockTests"
Option Explicit

''' <summary>   Unit test. Asserts instantiating and disposing of the Winsock framework. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInitializeAndDispose() As Assert

    ' this is required to initialize Winsock.  It will only ran once.
    Winsock.Initialize
    
    Set TestInitializeAndDispose = Assert.IsTrue(Winsock.Initiated, "Winsock should be initiated when a socket is created")
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If
    
    Set TestInitializeAndDispose = Assert.IsFalse(Winsock.Disposed, "Winsock should not be disposed when a socket is created")
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If
    
    Set TestInitializeAndDispose = Assert.AreEqual(Winsock.SocketCount, 0, _
        "Winsock socket count should be 0 as no sockets are registered but is " & Str$(Winsock.SocketCount))
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If

    ' test disposing of Winsock.
    Winsock.Dispose
    
    Set TestInitializeAndDispose = Assert.IsFalse(Winsock.Initiated, "Winsock should no longer be initiated after the last socket was set to nothing")
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If
    
    Set TestInitializeAndDispose = Assert.IsTrue(Winsock.Disposed, "Winsock should be disposed after the last socket was set to nothing")
    
    If Not TestInitializeAndDispose.AssertSuccessful Then
        Winsock.Dispose
        Exit Function
    End If
    
    Winsock.Dispose
    
End Function

''' <summary>   Unit test. Asserts getting the error description from the Windows API. </summary>
''' <returns>   An <see cref="Assert"/> instance of <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestGettingLastErrorDescription() As Assert

    Dim p_errorNumber As Long: p_errorNumber = 5
    Dim p_expected As String: p_expected = "Access is denied."
    Dim p_actual As String: p_actual = Winsock.LastErrorDescription(p_errorNumber)
    
    Set TestGettingLastErrorDescription = Assert.AreEqual(p_expected, p_actual, _
            "Winsock should get the correct error description for error number " & CStr(p_errorNumber))
    
End Function













