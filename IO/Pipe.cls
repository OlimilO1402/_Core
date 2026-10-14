VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PIPE_ACCESS_INBOUND  As Long = &H1 ' The flow of data in the pipe goes from client to server only. This mode gives the server the equivalent of GENERIC_READ access to the pipe. The client must specify GENERIC_WRITE access when connecting to the pipe. If the client must read pipe settings by calling the GetNamedPipeInfo or GetNamedPipeHandleState functions, the client must specify GENERIC_WRITE and FILE_READ_ATTRIBUTES access when connecting to the pipe.
Private Const PIPE_ACCESS_OUTBOUND As Long = &H2 ' The flow of data in the pipe goes from server to client only. This mode gives the server the equivalent of GENERIC_WRITE access to the pipe. The client must specify GENERIC_READ access when connecting to the pipe. If the client must change pipe settings by calling the SetNamedPipeHandleState function, the client must specify GENERIC_READ and FILE_WRITE_ATTRIBUTES access when connecting to the pipe.
Private Const PIPE_ACCESS_DUPLEX   As Long = &H3 ' The pipe is bi-directional; both server and client processes can read from and write to the pipe. This mode gives the server the equivalent of GENERIC_READ and GENERIC_WRITE access to the pipe. The client can specify GENERIC_READ or GENERIC_WRITE, or both, when it connects to the pipe using the CreateFile function.

Public Enum EPipeAccess
    PipeAccess_INBOUND = 1
    PipeAccess_OUTBOUND = 2
    PipeAccess_Duplex = 3
End Enum

Private Const FILE_FLAG_FIRST_PIPE_INSTANCE As Long = &H80000    ' If you attempt to create multiple instances of a pipe with this flag, creation of the first instance succeeds, but creation of the next instance fails with ERROR_ACCESS_DENIED.
Private Const FILE_FLAG_WRITE_THROUGH       As Long = &H80000000 ' Write-through mode is enabled. This mode affects only write operations on byte-type pipes and, then, only when the client and server processes are on different computers. If this mode is enabled, functions writing to a named pipe do not return until the data written is transmitted across the network and is in the pipe's buffer on the remote computer. If this mode is not enabled, the system enhances the efficiency of network operations by buffering data until a minimum number of bytes accumulate or until a maximum time elapses.
Private Const FILE_FLAG_OVERLAPPED          As Long = &H40000000 ' Overlapped mode is enabled. If this mode is enabled, functions performing read, write, and connect operations that may take a significant time to be completed can return immediately. This mode enables the thread that started the operation to perform other operations while the time-consuming operation executes in the background. For example, in overlapped mode, a thread can handle simultaneous input and output (I/O) operations on multiple instances of a pipe or perform simultaneous read and write operations on the same pipe handle. If overlapped mode is not enabled, functions performing read, write, and connect operations on the pipe handle do not return until the operation is finished. The ReadFileEx and WriteFileEx functions can only be used with a pipe handle in overlapped mode. The ReadFile, WriteFile, ConnectNamedPipe, and TransactNamedPipe functions can execute either synchronously or as overlapped operations.

Private Const WRITE_DAC                     As Long = &H40000    ' The caller will have write access to the named pipe's discretionary access control list (ACL).
Private Const WRITE_OWNER                   As Long = &H80000    ' The caller will have write access to the named pipe's owner.
Private Const ACCESS_SYSTEM_SECURITY        As Long = &H1000000  ' The caller will have write access to the named pipe's SACL. For more information, see Access-Control Lists (ACLs) and SACL Access Right.

Private Const PIPE_TYPE_BYTE                As Long = &H0        ' Data is written to the pipe as a stream of bytes. This mode cannot be used with PIPE_READMODE_MESSAGE. The pipe does not distinguish bytes written during different write operations.
Private Const PIPE_TYPE_MESSAGE             As Long = &H4        ' Data is written to the pipe as a stream of messages. The pipe treats the bytes written during each write operation as a message unit. The GetLastError function returns ERROR_MORE_DATA when a message is not read completely. This mode can be used with either PIPE_READMODE_MESSAGE or PIPE_READMODE_BYTE.

Public Enum EPipeType
    PipeType_BYTE = 0
    PipeType_MESSAGE = 4
End Enum

Private Const PIPE_READMODE_BYTE            As Long = &H0        ' Data is read from the pipe as a stream of bytes. This mode can be used with either PIPE_TYPE_MESSAGE or PIPE_TYPE_BYTE.
Private Const PIPE_READMODE_MESSAGE         As Long = &H2        ' Data is read from the pipe as a stream of messages. This mode can be only used if PIPE_TYPE_MESSAGE is also specified.

Private Const PIPE_WAIT                     As Long = &H0        ' Blocking mode is enabled. When the pipe handle is specified in the ReadFile, WriteFile, or ConnectNamedPipe function, the operations are not completed until there is data to read, all data is written, or a client is connected. Use of this mode can mean waiting indefinitely in some situations for a client process to perform an action.
Private Const PIPE_NOWAIT                   As Long = &H1        ' Nonblocking mode is enabled. In this mode, ReadFile, WriteFile, and ConnectNamedPipe always return immediately.
                                                                 ' Note that nonblocking mode is supported for compatibility with Microsoft LAN Manager version 2.0 and should not be used to achieve asynchronous I/O with named pipes. For more information on asynchronous pipe I/O, see Synchronous and Overlapped Input and Output.
Private Const PIPE_ACCEPT_REMOTE_CLIENTS    As Long = &H0        ' Connections from remote clients can be accepted and checked against the security descriptor for the pipe.
Private Const PIPE_REJECT_REMOTE_CLIENTS    As Long = &H8        ' Connections from remote clients are automatically rejected.

'Flags for GetNamedPipeInfo:
Private Const PIPE_CLIENT_END               As Long = &H0        ' The handle refers to the client end of a named pipe instance. This is the default.
Private Const PIPE_SERVER_END               As Long = &H1        ' The handle refers to the server end of a named pipe instance. If this value is not specified, the handle refers to the client end of a named pipe instance.
'Private Const PIPE_TYPE_BYTE                As Long = &H0        ' The named pipe is a byte pipe. This is the default.
'Private Const PIPE_TYPE_MESSAGE             As Long = &H4        ' The named pipe is a message pipe. If this value is not specified, the pipe is a byte pipe.

'State for GetNamedPipeHandleStateW
'Private Const PIPE_NOWAIT                   As Long = &H1        ' The pipe handle is in nonblocking mode. If this flag is not specified, the pipe handle is in blocking mode.
'Private Const PIPE_READMODE_MESSAGE         As Long = &H2        ' The pipe handle is in message-read mode. If this flag is not specified, the pipe handle is in byte-read mode.

'for CallNamedPipeW nTimeOut:
Private Const NMPWAIT_USE_DEFAULT_WAIT      As Long = &H0        ' Uses the default time-out specified in a call to the CreateNamedPipe function.
Private Const NMPWAIT_NOWAIT                As Long = &H1        ' Does not wait for the named pipe. If the named pipe is not available, the function returns an error.
Private Const NMPWAIT_WAIT_FOREVER          As Long = &HFFFFFFFF ' Waits indefinitely.

Private Const INVALID_HANDLE_VALUE          As Long = -1 '((Handle) - 1)
Private Const ERROR_INVALID_PARAMETER       As Long = 87
Private Const PIPE_UNLIMITED_INSTANCES      As Long = 255
Private Const NMPWAIT_USE_DEFAULT_WAIT      As Long = &H0

Private Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1     As Byte
    Control  As Long
    Owner    As Long
    Group    As Long
    Sacl     As LongPtr 'ACL
    Dacl     As LongPtr 'ACL
End Type

Private Type SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle       As Long
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-overlapped
'typedef struct _OVERLAPPED {
'  ULONG_PTR Internal;
'  ULONG_PTR InternalHigh;
'  union {
'    struct {
'      DWORD Offset;
'      DWORD OffsetHigh;
'    } DUMMYSTRUCTNAME;
'    PVOID Pointer;
'  } DUMMYUNIONNAME;
'  HANDLE    hEvent;
'} OVERLAPPED, *LPOVERLAPPED;
Private Type OVERLAPPED
    Internal     As LongPtr
    InternalHigh As LongPtr
    Offset       As Long
    OffsetHigh   As Long
    hEvent       As LongPtr
End Type
Private Type OVERLAPPEDP
    Internal     As LongPtr
    InternalHigh As LongPtr
    Pointer      As LongPtr
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-createnamedpipew
Private Declare Function CreateNamedPipeW Lib "kernel32" (ByVal lpName As LongPtr, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, _
ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, ByVal lpSecurityAttributes As LongPtr) As LongPtr
'If the function succeeds, the return value is a handle to the server end of a named pipe instance. 'If the function fails, the return value is INVALID_HANDLE_VALUE. To get extended error information, call GetLastError.
'nMaxInstances = 1..255

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-createpipe
Private Declare Function CreatePipe Lib "kernel32" (ByRef hReadPipe As LongPtr, ByRef hWritePipe As LongPtr, ByVal lpSecurityAttributes As LongPtr, ByVal nSize As Long) As Long 'BOOL
'If the function succeeds, the return value is nonzero. 'If the function fails, the return value is zero. To get extended error information, call GetLastError.

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-callnamedpipew
Private Declare Function CallNamedPipeW Lib "kernel32" (ByVal lpNamedPipeName As LongPtr, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesRead As Long, ByVal nTimeOut As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-connectnamedpipe
Private Declare Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, ByRef lpOverlapped As OVERLAPPED) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-disconnectnamedpipe
Private Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hPipe As LongPtr) As Long
'If the function succeeds, the return value is nonzero. 'If the function fails, the return value is zero. To get extended error information, call GetLastError.

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-getnamedpipeclientcomputernamew
Private Declare Function GetNamedPipeClientComputerNameW Lib "kernel32" (ByVal hNamedPipe As LongPtr, ByVal ClientComputerName As LongPtr, ByVal ClientComputerNameLength As Long) As Long
'Remarks: Windows 10, version 1709:  Pipes are only supported within an app-container; ie, from one UWP process to another UWP process that's part of the same app. Also, named pipes must use the syntax \\.\pipe\LOCAL\ for the pipe name.

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-getnamedpipehandlestatew
Private Declare Function GetNamedPipeHandleStateW Lib "kernel32" (ByVal hNamedPipe As LongPtr, ByRef lpState As LongPtr, ByRef lpCurInstances As Long, ByRef lpMaxCollectionCount As Long, ByRef lpCollectDataTimeout As Long, ByVal lpUserName As String, ByVal nMaxUserNameSize As Long) As Long
'If the function succeeds, the return value is nonzero. ' If the function fails, the return value is zero. To get extended error information, call GetLastError.

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-setnamedpipehandlestate
Private Declare Function SetNamedPipeHandleState Lib "kernel32.dll" (ByVal hNamedPipe As LongPtr, ByRef lpMode As Long, ByRef lpMaxCollectionCount As Long, ByRef lpCollectDataTimeout As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-getnamedpipeinfo
Private Declare Function GetNamedPipeInfo Lib "kernel32.dll" (ByVal hNamedPipe As LongPtr, ByRef lpFlags As Long, ByRef lpOutBufferSize As Long, ByRef lpInBufferSize As Long, ByRef lpMaxInstances As Long) As Long
'If the function succeeds, the return value is nonzero. 'If the function fails, the return value is zero. To get extended error information, call GetLastError.

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-peeknamedpipe
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
'If the function succeeds, the return value is nonzero. 'If the function fails, the return value is zero. To get extended error information, call GetLastError.

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-transactnamedpipe
Private Declare Function TransactNamedPipe Lib "kernel32.dll" (ByVal hNamedPipe As LongPtr, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesRead As Long, ByRef lpOverlapped As OVERLAPPED) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-waitnamedpipew
Private Declare Function WaitNamedPipeW Lib "kernel32.dll" (ByVal lpNamedPipeName As LongPtr, ByVal nTimeOut As Long) As Long


Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long


'Private iPipeFile As Long


Private m_SAttributes As SECURITY_ATTRIBUTES
Private m_SDescriptor As SECURITY_DESCRIPTOR
Private Const PathPipe  As String = "\\.\pipe\"
Private Const PathLocal As String = "LOCAL\"
Private m_Name    As String
Private m_Handle  As LongPtr 'of the named pipe
Private m_hRead   As LongPtr 'of the anonymouse pipe
Private m_hWrite  As LongPtr 'of the anonymouse pipe
Private m_IsLocal As Boolean

Private Sub Class_Initialize()
    m_SAttributes.lpSecurityDescriptor = VarPtr(m_SDescriptor)
End Sub

Private Sub Class_Terminate()
    Dim hr As Long
    If m_Handle Then
        hr = DisconnectNamedPipe(m_Handle)
        hr = CloseHandle(m_Handle)
        m_Handle = 0
    End If
    If m_hRead Then hr = CloseHandle(m_hRead):   m_hRead = 0
    If m_hWrite Then hr = CloseHandle(m_hWrite): m_hWrite = 0
End Sub

Friend Sub New_(Name As String, bIsLocal As Boolean, PipeAccess As EPipeAccess, PipeType As EPipeType, nInstances As Long)
    m_IsLocal = bIsLocal
    Dim hr As Long
    If Len(Name) Then
        If LCase(Left(Name, Len(PathPipe))) = PathPipe Then Name = Mid(Name, Len(PathPipe))
        If UCase(Left(Name, Len(PathLocal))) = PathLocal Then Name = Mid(Name, Len(PathLocal)): m_IsLocal = True
        m_Name = Name
        m_Handle = CreateNamedPipeW(StrPtr(Me.FullName), PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE, PIPE_UNLIMITED_INSTANCES, -1, -1, 0, VarPtr(m_SAttribs))
        If CLng(m_Handle) = INVALID_HANDLE_VALUE Then m_Handle = 0
    Else
        hr = CreatePipe(m_hRead, m_hWrite, VarPtr(m_SAttribs), 0)
    End If
End Sub

'Copy To MNew:
'Public Function Pipe(Optional ByVal Name As String, Optional bIsLocal As Boolean = False, _
'                     Optional ByVal PipeAccess As EPipeAccess = PipeAccess_Duplex, Optional ByVal PipeType As EPipeType = PipeType_BYTE, _
'                     Optional ByVal nInstances As Long = 255) As Pipe
'    Set Pipe = New Pipe: Pipe.New_ Name, bIsLocal, PipeAccess, PipeType, nInstances
'End Function

Public Property Get FullName() As String
    Dim p As String: p = PathPipe & IIf(m_IsLocal, PathLocal, "")
    FullName = p & m_Name
End Property
Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Get Handle() As LongPtr
    If m_Handle Then Handle = m_Handle Else Handle = m_hRead
End Property
Public Property Get hRead() As LongPtr
    hRead = m_hRead
End Property
Public Property Get hWrite() As LongPtr
    hWrite = m_hWrite
End Property

Public Sub CallPipe()
    
End Sub
