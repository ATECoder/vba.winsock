VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Winsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Winsock initialization and termination. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_

    ' true if Winsock was initiated
    Initiated As Boolean

    ' true if Winsock was disposed
    Disposed As Boolean

    ' counts the number of sockets
    SocketCount As Integer

End Type

Private This As this_

''' <summary>   Initiates use of the Winsock DLL by a process. </summary>
''' <param name="a_version">   [Long] Specifies the Winsock version. </param>
''' <param name="a_wsa">       [ws32_WSADATA] A pointer to the ws32_WSADATA data structure that is to receive
'''                            details of the Windows Sockets implementation. </param>
''' <returns>   If successful, the WSAStartup function returns zero. Otherwise, it returns one of
''' the error codes listed below. The WSAStartup function directly returns the extended error code
''' in the return value for this function. A call to the WSAGetLastError function is not needed and should not be used.
''' </returns>
Private Declare PtrSafe Function WSAStartup Lib "wsock32.dll" ( _
        ByVal a_version As Long, ByRef a_wsa As wsock32.ws32_WSADATA) As Long

''' <summary>   Returns the error status for the last Windows Sockets operation that failed. </summary>
''' <remarks>   WSAGetLastError should be called immediately. This is necessary because some functions may reset
''' the last extended error code to 0 if they succeed. </remarks>
''' <returns>   The return value indicates the error code for this thread's last Windows Sockets
''' operation that failed. <returns>
Private Declare PtrSafe Function WSAGetLastError Lib "wsock32.dll" () As Long

''' <summary>   terminates use of the Winsock dll. </summary>
''' <remarks>   In a multi-threaded environment, WSACleanup terminates Windows Sockets operations
'''   for all threads. </remarks>
''' <returns>
'''   The return value is zero if the operation was successful. Otherwise, the value
'''   ws32_SOCKET_ERROR is returned, and a specific error number can be retrieved by calling WSAGetLastError.
''' <returns>
Private Declare PtrSafe Function WSACleanup Lib "wsock32.dll" () As Long

Private Declare PtrSafe Function FormatMessage_ Lib "kernel32" Alias "FormatMessageA" ( _
         ByVal dwFlags As Long, _
               lpSource As Any, _
         ByVal dwMessageId As Long, _
         ByVal dwLanguageId As Long, _
         ByVal lpBuffer As String, _
         ByVal nSize As Long, _
               a_arguments As LongPtr) As Long
               
Private Declare PtrSafe Function GetSystemDefaultLangID_ Lib "kernel32" Alias "GetSystemDefaultLangID" () As Integer
               
''' <summary>   Initializes Winsock. </summary>
'''< remarks>   This is a singleton class.
''' <see href="https://stackoverflow.com/questions/396117/class-static-methods-in-vba"/>
''' <see href="https://web.archive.org/web/20160418020100/http://www.utteraccess.com/wiki/Index.php/Singleton"/>
''' This method is called on the first method (not property) access of this class.
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Initializes Winsock. </summary>
''' <remarks>   This is a singleton class.
''' <see href="https://stackoverflow.com/questions/396117/class-static-methods-in-vba"/>
''' <see href="https://web.archive.org/web/20160418020100/http://www.utteraccess.com/wiki/Index.php/Singleton"/>
''' </remarks>
''' <returns>   [Long] The return value is zero if the operation was successful. Otherwise, the value
''' ws32_SOCKET_ERROR is returned, a specific error number can be retrieved by calling WSAGetLastError.
''' and the error is raised.
''' <returns>
Public Function Initialize() As Long

    Dim p_result As Long: p_result = 0

    ' skip if the class was already initialized.
    If This.Initiated Then
        Initialize = p_result
        Exit Function
    End If
    
    Dim p_startUpInfo As wsock32.ws32_WSADATA
    
    ' Initialize Winsock DLL
    
    p_result = WSAStartup(wsock32.ws32_VERSION, p_startUpInfo)
    
    If p_result <> 0 Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.WinsockStartingError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".Initialize", _
            " Failed initializing WINSOCK. winsock32.WSAStartup(wsock32.ws32_VERSION, p_startUpInfo) returned " & _
            VBA.CStr(p_result) & Winsock.LastErrorMessage()
    End If
    This.Initiated = (p_result = 0)
    This.Disposed = False
    This.SocketCount = 0
    Initialize = p_result

End Function

''' <summary>   Registers this socket with the Winsock class. </summary>
''' <remarks>   Call this method to register the socket with Winsock thus managing when
''' Winsock can be disposed.</remarks>
''' <param name="a_activeSocket">   The socket that is registering. </param>
Public Sub Register(ByVal a_activeSocket As IPv4StreamSocket)
    
    If IsObject(a_activeSocket) Then
        This.SocketCount = This.SocketCount + 1
    End If

End Sub

''' <summary>   Unregisters this socket from the Winsock class. </summary>
''' <remarks>   Call this method to remove the socket from the class so it can be disposed when having
''' no registered sockets.</remarks>
''' <param name="activeSocket">   The socket that is unregistering. </param>
Public Sub Unregister(ByVal a_activeSocket As IPv4StreamSocket)
    
    If IsObject(a_activeSocket) And (Me.SocketCount > 0) Then
        This.SocketCount = This.SocketCount - 1
        If Me.SocketCount = 0 Then
            Me.Dispose
        End If
    End If

End Sub

''' <summary>   Returns true if Winsock was initiated. </summary>
Public Property Get Initiated() As Boolean
    
    Initiated = This.Initiated

End Property

''' <summary>   Returns true if Winsock was Disposed. </summary>
Public Property Get Disposed() As Boolean
    
    Disposed = This.Disposed

End Property

''' <summary>   Builds the last error message including the error number. </summary>
''' <param name="a_prefix">   [Optional, String, '; Last WINSOCK error: '] the prefix for the message. </param>
''' <param name="a_suffix">   [Optional, String, '."] the suffix for the message. </param>
''' <returns>   [String]. <returns>
Public Function LastErrorMessage(Optional ByVal a_prefix As String = "; Last WINSOCK error: ", _
    Optional ByVal a_suffix As String = ".") As String
    
    Dim p_reply As String: p_reply = VBA.vbNullString
    Dim p_number As Long: p_number = Me.LastError
    If p_number <> 0 Then
        p_reply = a_prefix
        Dim p_description As String
        p_description = Me.LastErrorDescription(p_number)
        If VBA.vbNullString <> p_description Then
            p_reply = p_reply & p_description & " "
        End If
        p_reply = p_reply & "#" & VBA.CStr(p_number)
    End If
    LastErrorMessage = p_reply & a_suffix

End Function

''' <summary>   returns the last Winsock 32 error . </summary>
Public Property Get LastError() As Long
    
    LastError = WSAGetLastError()

End Property

''' <summary>   Gets the last windows base error message for the specified error number. </summary>
''' <returns>   [String]. <returns>
Public Function LastErrorDescription(ByVal a_errorNumber As Long) As String

    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
    Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
    Const FORMAT_MESSAGE_TEXT_LEN As Long = &HA0
    
    ' The language into which to format the message. 0 = default language.
    
    Dim p_languageId As Long: p_languageId = 0
    
    Dim p_errMessage As String * FORMAT_MESSAGE_TEXT_LEN
    Dim p_result As Long
    p_result = FormatMessage_(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
           0, _
           a_errorNumber, _
           p_languageId, _
           p_errMessage, _
           FORMAT_MESSAGE_TEXT_LEN, _
           0)
    
    If p_result = 0 Then
        LastErrorDescription = vbNullString
    Else
        p_errMessage = Split(p_errMessage, vbCrLf)(0)
        LastErrorDescription = Trim(p_errMessage)
    End If
           
End Function

''' <summary>   Terminates use of the Winsock dll. </summary>
''' <returns>
''' The return value is zero if the operation was successful. Otherwise, the value
''' ws32_SOCKET_ERROR is returned, a specific error number can be retrieved by calling WSAGetLastError.
''' and the error is raised.
''' <returns>
Public Function Cleanup() As Long
    
    ' Shutdown Winsock DLL
    Dim p_result As Long
    
    p_result = WSACleanup()
    If p_result = wsock32.ws32_SOCKET_ERROR Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.WinsockCleanupError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".Cleanup", _
            " Failed cleaning up WINSOCK. winsock32.WSACleanup() returned " & _
            VBA.CStr(p_result) & Winsock.LastErrorMessage()
    End If
    Cleanup = p_result

End Function

''' <summary>   Returns the number of sockets associated with this Winsock instance. </summary>
Public Property Get SocketCount() As Integer
    SocketCount = This.SocketCount
End Property

''' <summary>   Tries to parse the TCP address in the form of <c>host:port</c>. </summary>
''' <param name="a_address">   [String] An IPv4 dotted-decimal <c>host:port</c> address. </param>
''' <param name="a_host">      [Out, String] An IPv4 dotted-decimal host address. </param>
''' <param name="a_port">      [Out, Long] The port that the server is listening on. </param>
''' <param name="a_host">      [Out, String] Detaisl the failure reason. </param>
''' <returns>   [Boolean] True if the address parsed. </returns>
Public Function TryParseAddress(ByVal a_address As String, ByRef a_host As String, ByRef a_port As Long, _
    ByRef a_details As String) As Boolean

    Dim p_result As Boolean
    
    If VBA.vbNullString = a_address Then
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".TryParseAddress" & _
            " 'a_address' argument '" & a_address & "' is empty."
        p_result = False
    Else
    
        a_host = cc_isr_Core.StringExtensions.Pop(a_address, ":")
        
        If VBA.vbNullString = a_host Then
            
            a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".TryParseAddress" & _
                " Host part in '" & a_address & "' is empty."
            p_result = False
        
        Else
            
            If 1 = VBA.InStr(1, a_host, "localhost", vbTextCompare) Then: a_host = "127.0.0.1"
            
            Dim p_port As String
            p_port = cc_isr_Core.StringExtensions.Pop(a_address, ":")
                
            If VBA.vbNullString = p_port Then
            
                a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".TryParseAddress" & _
                    " Port part in '" & a_address & "' is empty."
                p_result = False
            
            Else
                
                p_result = cc_isr_Core.StringExtensions.TryParseLong(p_port, a_port, a_details)
            
            End If
            
        End If
    
    End If
    TryParseAddress = p_result
    
End Function

''' <summary>   Terminates Winsock if all sockets released their association with this instance. </summary>
Public Sub Dispose()
    
    On Error Resume Next
    
    If This.Initiated And Not This.Disposed And (This.SocketCount = 0) Then
        This.Disposed = True
        This.Initiated = False
        
        ' Shutdown Winsock DLL
        Dim p_result As Long: p_result = WSACleanup()
        If p_result = wsock32.ws32_SOCKET_ERROR Then
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.WinsockCleanupError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".Dispose", _
                " Failed cleaning up WINSOCK. winsock32.WSACleanup() returned " & _
                VBA.CStr(p_result) & Winsock.LastErrorMessage()
        End If
    End If
    
    On Error GoTo 0

End Sub

