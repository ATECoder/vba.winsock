Attribute VB_Name = "IPv4StreamSocketTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   IPV4 Stream socket tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts creating a socket. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCreateSocket() As Assert

    Dim p_outcome As Assert
    
    Dim p_socket As New IPv4StreamSocket
    
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
    
    Debug.Print p_outcome.BuildReport("TestCreateSocket")
    
    Set TestCreateSocket = p_outcome
    
End Function




