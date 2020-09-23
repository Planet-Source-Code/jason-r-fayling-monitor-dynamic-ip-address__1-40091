VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGetIPAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private Const MIN_SOCKETS_REQD As Long = 1
Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Private Const SOCKET_ERROR As Long = -1
Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129
Private Const MAX_WSADescription = 256
Const MAX_WSASYSStatus = 128

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

Private Type WSADataInfo
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Event ErrorOccured(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)

Private Declare Function WSAStartupInfo Lib "WSOCK32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSADATA As WSADataInfo) As Long
Private Declare Function WSACleanup Lib "WSOCK32" () As Long
Private Declare Function WSAGetLastError Lib "WSOCK32" () As Long
Private Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Private Declare Function GetHostName Lib "WSOCK32" Alias "gethostname" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32" (ByVal szHost As String) As Long
Private Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Function GetIPAddress(Optional ByVal sHostName As String) As String

On Error GoTo ErrorHandle

Dim lpHost As Long
Dim HOST As HOSTENT
Dim dwIPAddr As Long
Dim tmpIPAddr() As Byte
Dim i As Integer
Dim sIPAddr As String
    
    If Not SocketsInitialize() Then
        GetIPAddress = ""
        GoTo ClearVariables
    End If
    
    If Len(sHostName) = 0 Then
        sHostName = String(256, " ")
    
        If GetHostName(sHostName, Len(sHostName)) = SOCKET_ERROR Then
            GetIPAddress = ""
            mErrorHandle 90001, "GetIPAddress", "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
            GoTo ClearVariables
        End If
    
    End If
    
    
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)
    If lpHost = 0 Then
        GetIPAddress = ""
        mErrorHandle 90002, "GetIPAddress", "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
        GoTo ClearVariables
    End If
    
    CopyMemoryIP HOST, lpHost, Len(HOST)
    CopyMemoryIP dwIPAddr, HOST.hAddrList, 4
    
    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemoryIP tmpIPAddr(1), dwIPAddr, HOST.hLen
    For i = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    
ClearVariables:
    SocketsCleanup
    Exit Function
    
ErrorHandle:
    mErrorHandle Err.Number, "GetIPAddress", Err.Description
    GoTo ClearVariables
    
End Function
Private Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function
Private Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function
Private Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        mErrorHandle 90003, "SocketCleanup", "Socket error occurred in Cleanup."
    End If
End Sub
Private Function SocketsInitialize() As Boolean

On Error GoTo ErrorHandle

Dim WSAD As WSAData
Dim sLoByte As String
Dim sHiByte As String

    SocketsInitialize = False
    
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        mErrorHandle 90004, "SocketsInitialize", "The 32-bit Windows Socket is not responding."
        GoTo ClearVariables
    End If
    
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        mErrorHandle 90005, "SocketsInitialize", "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
        GoTo ClearVariables
    End If
    
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
        mErrorHandle 90006, "SocketsInitialize", "Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."
        Exit Function
    End If
    
    'must be OK, so lets do it
    SocketsInitialize = True

ClearVariables:
    Exit Function
    
ErrorHandle:
    mErrorHandle Err.Number, "SocketsInitialize", Err.Description
    GoTo ClearVariables
    
End Function
Private Sub mErrorHandle(ByVal lErrorNumber As Long, ByVal sErrorSource As String, ByVal sErrorDescription As String)
    RaiseEvent ErrorOccured(lErrorNumber, "CGetIPAddress." & sErrorSource, sErrorDescription)
End Sub
