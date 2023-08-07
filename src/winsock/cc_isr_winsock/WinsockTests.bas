Attribute VB_Name = "WinsockTests"
Option Explicit

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
    
    Debug.Print p_outcome.BuildReport("TestGettingLastErrorDescription")
    
    Set TestGettingLastErrorDescription = p_outcome
    
End Function













