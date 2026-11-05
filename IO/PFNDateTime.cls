VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PFNDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#If VBA7 = 0 Then
'    Private Enum LongPtr
'        [_]
'    End Enum
'#End If

Private Const GENERIC_READ               As Long = &H80000000
Private Const GENERIC_WRITE              As Long = &H40000000
Private Const FILE_SHARE_READ            As Long = &H1
Private Const FILE_SHARE_WRITE           As Long = &H2
Private Const OPEN_EXISTING              As Long = 3
Private Const FILE_FLAG_BACKUP_SEMANTICS As Long = &H2000000

Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFilename As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As LongPtr
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private m_PFN    As PathFileName
Private m_Handle As LongPtr
Private m_CreationTime   As Date
Private m_LastAccessTime As Date
Private m_LastWriteTime  As Date

Private Sub Class_Initialize()
    CClose
End Sub

Private Sub Class_Terminate()
    CClose
End Sub

Friend Sub New_(aPFN As PathFileName)
    Set m_PFN = aPFN
    m_Handle = CreateFileW(StrPtr(m_PFN.Value), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
    m_CreationTime = GetCreationTime
    m_LastWriteTime = GetLastWriteTime
    m_LastAccessTime = GetLastAccessTime
End Sub

Public Sub NewC(other As PFNDateTime)
    Me.New_ m_PFN
    m_CreationTime = other.CreationTime
    m_LastAccessTime = other.LastAccessTime
    m_LastWriteTime = other.LastWriteTime
End Sub

Public Function Clone() As PFNDateTime
    Set Clone = New PFNDateTime: Clone.NewC Me
End Function

Public Sub CClose()
    If m_Handle <= 0 Then Exit Sub
    CloseHandle m_Handle
    m_Handle = 0
End Sub

'Public Property Get Handle() As LongPtr
'    Handle = m_Handle
'End Property

Public Property Get IsEqual(other As PFNDateTime) As Boolean
    'here only compare the datetimes not the filename
    If Not (Me.CreationTime = other.CreationTime) Then Exit Property
    If Not (Me.LastAccessTime = other.LastAccessTime) Then Exit Property
    If Not (Me.LastWriteTime = other.LastWriteTime) Then Exit Property
    IsEqual = True
End Property

Public Property Get PathFileName() As PathFileName
    Set PathFileName = m_PFN
End Property

Private Function GetCreationTime() As Date
    If m_Handle <= 0 Then Exit Function
    Dim ct As FILETIME, lat As FILETIME, lwt As FILETIME
    If GetFileTime(m_Handle, ct, lat, lwt) = 0 Then
        MsgBox "Could not get FileTime from file: " & m_PFN.Value
        Exit Function
    End If
    GetCreationTime = MTime.SystemTime_ToDate(MTime.SystemTime_ToTzSpecificLocalTime(MTime.FileTime_ToSystemTime(ct)))
End Function
Public Property Get CreationTime() As Date
    CreationTime = m_CreationTime
End Property
Public Property Let CreationTime(ByVal aDateTime As Date)
    m_CreationTime = aDateTime
    If m_Handle <= 0 Then Exit Property
    Dim ct As FILETIME, lat As FILETIME, lwt As FILETIME
    If GetFileTime(m_Handle, ct, lat, lwt) = 0 Then
        MsgBox "Could not get FileTime from file: " & m_PFN.Value
        Exit Property
    End If
    ct = MTime.Date_ToFileTime(m_CreationTime)
    ct = MTime.LocalFileTime_ToFileTime(ct)
    If SetFileTime(m_Handle, ct, lat, lwt) = 0 Then
        MsgBox "Could not write FileTime to file: " & m_PFN.Value
        Exit Property
    End If
End Property

Private Function GetLastAccessTime() As Date
    If m_Handle <= 0 Then Exit Function
    Dim ct As FILETIME, lat As FILETIME, lwt As FILETIME
    If GetFileTime(m_Handle, ct, lat, lwt) = 0 Then
        MsgBox "Could not get FileTime from file: " & m_PFN.Value
        Exit Function
    End If
    GetLastAccessTime = MTime.SystemTime_ToDate(MTime.SystemTime_ToTzSpecificLocalTime(MTime.FileTime_ToSystemTime(lat)))
End Function
Public Property Get LastAccessTime() As Date
    LastAccessTime = m_LastAccessTime
End Property
Public Property Let LastAccessTime(ByVal aDateTime As Date)
    m_LastAccessTime = aDateTime
    If m_Handle <= 0 Then Exit Property
    Dim ct As FILETIME, lat As FILETIME, lwt As FILETIME
    If GetFileTime(m_Handle, ct, lat, lwt) = 0 Then
        MsgBox "Could not get FileTime from file: " & vbCrLf & m_PFN.Value
        Exit Property
    End If
    lat = MTime.Date_ToFileTime(m_LastAccessTime)
    lat = MTime.LocalFileTime_ToFileTime(lat)
    If SetFileTime(m_Handle, ct, lat, lwt) = 0 Then
        MsgBox "Could not write FileTime to file: " & vbCrLf & m_PFN.Value
        Exit Property
    End If
End Property

Private Function GetLastWriteTime() As Date
    If m_Handle <= 0 Then Exit Function
    Dim ct As FILETIME, lat As FILETIME, lwt As FILETIME
    If GetFileTime(m_Handle, ct, lat, lwt) = 0 Then
        MsgBox "Could not get FileTime from file: " & vbCrLf & m_PFN.Value
        Exit Function
    End If
    GetLastWriteTime = MTime.SystemTime_ToDate(MTime.SystemTime_ToTzSpecificLocalTime(MTime.FileTime_ToSystemTime(lwt)))
End Function
Public Property Get LastWriteTime() As Date
    LastWriteTime = m_LastWriteTime
End Property
Public Property Let LastWriteTime(ByVal aDateTime As Date)
    m_LastWriteTime = aDateTime
    If m_Handle <= 0 Then Exit Property
    Dim ct As FILETIME, lat As FILETIME, lwt As FILETIME
    If GetFileTime(m_Handle, ct, lat, lwt) = 0 Then
        MsgBox "Could not get FileTime from file: " & vbCrLf & m_PFN.Value
        Exit Property
    End If
    lwt = MTime.Date_ToFileTime(m_LastWriteTime)
    lwt = MTime.LocalFileTime_ToFileTime(lwt)
    If SetFileTime(m_Handle, ct, lat, lwt) = 0 Then
        MsgBox "Could not write FileTime to file: " & vbCrLf & m_PFN.Value
        Exit Property
    End If
End Property
