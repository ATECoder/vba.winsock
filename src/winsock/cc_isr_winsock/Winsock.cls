VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Winsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' true if Winsock was initiated
Private m_initiated As Boolean

' true if Winsock was disposed
Private m_disposed As Boolean

' counts the number of sockets
Private m_socketCount As Integer

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
               Arguments As LongPtr) As Long
               
Private Declare PtrSafe Function GetSystemDefaultLangID_ Lib "kernel32" Alias "GetSystemDefaultLangID" () As Integer
               
''' <summary>   Initializes Winsock. </summary>
'''< remarks>   This is a singleton class.
''' <see href="https://stackoverflow.com/questions/396117/class-static-methods-in-vba"/>
''' <see href="https://web.archive.org/web/20160418020100/http://www.utteraccess.com/wiki/Index.php/Singleton"/>
''' This method is call on the first method (not property) access of this class.
''' </remarks>
Private Sub Class_Initialize()
    Debug.Print "initializing Winsock"
End Sub

''' <summary>   Initializes Winsock. </summary>
'''< remarks>   This is a singleton class.
''' <see href="https://stackoverflow.com/questions/396117/class-static-methods-in-vba"/>
''' <see href="https://web.archive.org/web/20160418020100/http://www.utteraccess.com/wiki/Index.php/Singleton"/>
''' <returns>  [Long] The return value is zero if the operation was successful. Otherwise, the value
''' ws32_SOCKET_ERROR is returned, a specific error number can be retrieved by calling WSAGetLastError.
''' and the error is raised.
''' <returns>
Public Function Initialize() As Long

    Dim p_result As Long: p_result = 0

    ' skip if the class was already initialized.
    If m_initiated Then
        Initialize = p_result
        Exit Function
    End If
    
    Dim p_startUpInfo As wsock32.ws32_WSADATA
    
    ' Initialize Winsock DLL
    
    p_result = WSAStartup(wsock32.ws32_VERSION, p_startUpInfo)
    
    If p_result <> 0 Then
        Err.Raise cc_isr_Core.UserDefinedErrors.WinsockStartingError.Code, "IdentityReader.Winsock.Initialize", _
                  cc_isr_Core.UserDefinedErrors.WinsockStartingError.Description & _
                  Me.LastErrorMessage(" Last error: ")
    End If
    m_initiated = (p_result = 0)
    m_disposed = False
    m_socketCount = 0
    Initialize = p_result

End Function

''' <summary>   Registers this socket with the Winsock class. </summary>
''' <remarks>   Call this method to register the socket with Winsock thus managing when
''' Winsock can be disposed.</remarks>
''' <param name="a_activeSocket">   The socket that is registering. </param>
Public Sub Register(ByVal a_activeSocket As IPv4StreamSocket)
    
    If IsObject(a_activeSocket) Then
        m_socketCount = m_socketCount + 1
    End If

End Sub

''' <summary>   Un-Registers this socket from the Winsock class. </summary>
''' <remarks>   Call this method to remove the socket from the class so it can be disposed when having
''' no registered sockets.</remarks>
''' <param name="activeSocket">   The socket that is unregistering. </param>
Public Sub UnRegister(ByVal a_activeSocket As IPv4StreamSocket)
    
    If IsObject(a_activeSocket) And (Me.SocketCount > 0) Then
        m_socketCount = m_socketCount - 1
        If Me.SocketCount = 0 Then
            Me.Dispose
        End If
    End If

End Sub

''' <summary>   Returns true if Winsock was initiated. </summary>
Public Property Get Initiated() As Boolean
    
    Initiated = m_initiated

End Property

''' <summary>   Returns true if Winsock was Disposed. </summary>
Public Property Get Disposed() As Boolean
    
    Disposed = m_disposed

End Property

''' <summary>   Builds the last error message including the error number. </summary>
''' <returns>   [String]. <returns>
Public Function LastErrorMessage(ByVal prefix As String) As String
    
    Dim p_number As Long: p_number = Me.LastError
    If p_number = 0 Then
        LastErrorMessage vbNullString
    Else
        Dim p_description As String
        p_description = Me.LastErrorDescription(p_number)
        If StringExtensions.IsNullOrEmpty(p_description) Then
            LastErrorMessage = prefix & p_description & " #" & CStr(p_number) & "."
        Else
            LastErrorMessage = prefix & "#" & CStr(p_number) & "."
        End If
    End If

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
    
    Dim langId As Long: langId = 0
    
    Dim p_errMessage As String * FORMAT_MESSAGE_TEXT_LEN
    Dim p_result As Long
    p_result = FormatMessage_(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
           0, _
           a_errorNumber, _
           langId, _
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
        Err.Raise cc_isr_Core.UserDefinedErrors.WinsockCleanupError.Code, "IdentityReader.Winsock.Cleanup", _
                  cc_isr_Core.UserDefinedErrors.WinsockCleanupError.Description & _
                  Winsock.LastErrorMessage(" Last error: ")
    End If
    Cleanup = p_result

End Function

''' <summary>   Returns the number of sockets associated with this Winsock instance. </summary>
Public Property Get SocketCount() As Integer
    SocketCount = m_socketCount
End Property

''' <summary>   Terminates Winsock if all sockets released their association with this instance. </summary>
Public Sub Dispose()
    
    On Error Resume Next
    
    If m_initiated And Not m_disposed And (m_socketCount = 0) Then
        m_disposed = True
        m_initiated = False
        
        ' Shutdown Winsock DLL
        Dim p_result As Long: p_result = WSACleanup()
        If p_result = wsock32.ws32_SOCKET_ERROR Then
            Err.Raise cc_isr_Core.UserDefinedErrors.WinsockCleanupError.Code, "IdentityReader.Winsock.Cleanup", _
                      cc_isr_Core.UserDefinedErrors.WinsockCleanupError.Description & _
                      Winsock.LastErrorMessage(" Last error: ")
        End If
    End If
    
    On Error GoTo 0

End Sub

