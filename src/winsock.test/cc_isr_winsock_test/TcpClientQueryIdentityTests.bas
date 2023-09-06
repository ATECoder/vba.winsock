Attribute VB_Name = "TcpClientQueryIdentityTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Client query identity Tests. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
    Host As String
    Port As Long
    PrologixPort As Long
    SocketReceiveTimeout As Integer
    Client As TcpClient
    DelayStopper As cc_isr_Core_IO.Stopwatch
    ErrTracer As IErrTracer
End Type

Private This As this_

Public Sub RunTest(ByVal a_testNumber As Integer)
    BeforeEach
    Select Case a_testNumber
        Case 1
            TestTcpClentShouldQueryIdentity
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
    This.Host = "192.168.0.252"
    This.Port = 1234
    This.PrologixPort = 1234
    This.SocketReceiveTimeout = 100
    
    Set This.BeforeAllAssert = Assert.IsTrue(True, "initialize the overall assert.")
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    Set This.DelayStopper = cc_isr_Core_IO.Factory.NewStopwatch
        
    Set This.ErrTracer = New ErrTracer
    
    Set This.Client = cc_isr_Winsock.Factory.NewTcpClient()
    
    ' trap errors in case connection fails rendering all tests inconclusive.
    
    On Error Resume Next
    
    This.Client.OpenConnection This.Host, This.Port
    
    Dim p_leftoverErrorMessage As String
    p_leftoverErrorMessage = VBA.vbNullString
    
    If Err.Number <> 0 Then
        p_leftoverErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        Set This.BeforeAllAssert = Assert.Inconclusive("IPV4 Stream Client failed to connect: " & _
            p_leftoverErrorMessage)
    ElseIf cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        Set This.BeforeAllAssert = Assert.Inconclusive("IPV4 Stream Client failed to connect: " & _
            p_leftoverErrorMessage)
    ElseIf This.Client.Connected Then
        Set This.BeforeAllAssert = Assert.IsTrue(True, "Connected")
    Else
        Set This.BeforeAllAssert = Assert.Inconclusive("IPV4 Stream Client should be connected")
    End If
    
    This.ErrTracer.TraceError p_leftoverErrorMessage
    
    ' clear the error object.
    On Error GoTo 0
    
End Sub

Public Sub BeforeEach()

    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
        
        Set This.BeforeEachAssert = IIf(This.Client.Connected, _
            Assert.IsTrue(True, "Connected"), _
            Assert.Inconclusive("IPV4 Stream Client should be connected"))
    
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
    
    Set This.BeforeEachAssert = Assert.AreEqual("1", p_reply, _
            "Operation completion should send the correct reply.")
                    
End Sub

Public Sub AfterEach()

    Dim p_command As String
    Dim p_sentCount As Integer
    Dim p_reply As String

    If This.BeforeEachAssert.AssertSuccessful Then
    
        ' clear errors if any so as to leave the instrument without errors.
        p_sentCount = This.Client.SendMessage("*CLS;*WAI;*OPC?" & VBA.vbLf)
        This.DelayStopper.Wait 5
        p_reply = This.Client.ReceiveRaw
        This.DelayStopper.Wait 5

        ' Restore Prologix device
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
    
    Set This.BeforeEachAssert = Nothing
        
End Sub

Public Sub AfterAll()
    
    ' disconnect if connected
    If Not This.Client Is Nothing Then _
        This.Client.CloseConnection

    Set This.Client = Nothing

    Set This.BeforeAllAssert = Nothing

End Sub

''' <summary>   Unit test. Asserts that the TCP Client should query a device identity. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTcpClentShouldQueryIdentity() As Assert

    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    Dim p_command As String: p_command = "*IDN?"
    Dim p_sentCount As Integer
    Dim p_identity As String
    
    If p_outcome.AssertSuccessful Then
            
        ' send the command
        p_sentCount = This.Client.SendMessage(p_command & VBA.vbLf)
        This.DelayStopper.Wait 5
            
        p_identity = This.Client.ReceiveRaw()

    End If

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestTcpClentShouldQueryIdentity")
    
    Set TestTcpClentShouldQueryIdentity = p_outcome
    
End Function

