VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperatingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public NotInheritable Class OperatingSystem
'     Member von: System
'Instancing: 6 - GlobalMultiUse
Option Explicit
'Implements ICloneable
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hCurProc As LongPtr, ByRef bValue As Long) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As Long

'Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As string) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_TABLETPC    As Long = 86
Private Const SM_MEDIACENTER As Long = 87
Private Const SM_STARTER     As Long = 88
Private Const SM_SERVERR2    As Long = 89

Public Enum PlatformID
    Win32S = 0
    Win32Windows = 1
    WinNT = 2
    WinCE = 3
    UNIX = 4
    
End Enum
Private Const VER_SUITE_PERSONAL As Long = &H200&
Private Const VER_NT_WORKSTATION As Long = &H1&

Private m_Platform     As PlatformID
Private m_Version      As Version
Private m_ServicePack  As String
Private m_wSuiteMask   As Long
'Private m_PlatFormID   As Long '??

'Public Sub New(ByVal platform As System.PlatformID, ByVal version As System.Version)
Public Sub New_(Platform As PlatformID, aVersion As Version) 'As OperatingSystem
    m_Platform = Platform
    Set m_Version = aVersion
End Sub

'Public Function OperatingSystem(ByVal Platform As PlatformID, ByVal SVersion As System.Version) As OperatingSystem
'  Set OperatingSystem = New OperatingSystem: OperatingSystem.New_ Platform, SVersion
'End Function

'################'  v     Implements ICloneable      v   '################'
'Private Function ICloneable_Clone() As Object
'  Set ICloneable_Clone = Clone
'End Function
'Public Overridable NotOverridable Function Clone() As Object
'Private Function Clone() As Object
'  Set Clone = Me
'End Function

'Public Overrides Function ToString() As String
Public Function ToStr() As String 'cString
    Dim s As String
    'Set ToString = New cString
    s = "Microsoft "
    Select Case m_Platform
    Case PlatformID.WinNT
        Select Case m_Version.major
        Case 10:
           Select Case m_Version.minor
           Case 0:
              s = s & "Windows 10 "
           Case Else
           End Select
        Case 6:
           Select Case m_Version.minor
           Case 0:
              s = s & "Windows Vista "
            'wie geht die Unterscheidung zwischen
            '"Windows Vista" und "Windows Server 2008 (LongHorn)" ?
           Case 1:
              s = s & "Windows 7 "
           End Select
            
        Case 5
            Select Case m_Version.minor
            Case 2:               s = s & "Windows Server 2003 "
                If GetSystemMetrics(SM_SERVERR2) > 0 Then
                                        s = s & "R2 "
                End If
            Case 1:               s = s & "Windows XP "
                If m_wSuiteMask And VER_SUITE_PERSONAL Then
                                        s = s & "Home Edition "
                End If
                If GetSystemMetrics(SM_MEDIACENTER) > 0 Then
                                        s = s & "Media Center Edition "
                ElseIf GetSystemMetrics(SM_STARTER) > 0 Then
                                        s = s & "Starter Edition "
                ElseIf GetSystemMetrics(SM_TABLETPC) > 0 Then
                                        s = s & "Tablet PC Edition "
                End If
            Case 0:               s = s & "Windows 2000 "
            End Select
        Case 4:                 s = s & "Windows NT "
        End Select
    Case PlatformID.Win32Windows
        Select Case m_Version.minor
        Case 0:                 s = s & "Windows 95 "
        Case 10:                s = s & "Windows 98 "
        Case 90:                s = s & "Windows ME "
        End Select
    Case PlatformID.WinCE:    s = s & "Windows CE "
    Case PlatformID.Win32S:   s = s & "Win32S "
    End Select
    ToStr = s & m_Version.ToStr & " " & BitToStr & " " & m_ServicePack
End Function

'Public ReadOnly Property Platform() As System.PlatformID
Public Property Get Platform() As PlatformID
    Platform = m_Platform
End Property

'Public ReadOnly Property Version() As System.Version
Public Property Get Version() As Version
    Set Version = m_Version
End Property

'Public ReadOnly Property ServicePack() As String
Public Property Get ServicePack() As String 'cString
    'Set ServicePack = New cString
    ServicePack = m_ServicePack
End Property

Friend Sub LetServicePack(ByVal StrVal As String)
    Dim np As Long
    If Len(StrVal) > 0 Then
        If Len(StrVal) > 1 Then
            np = InStr(1, StrVal, vbNullChar, vbTextCompare) - 1
            If np > 0 Then StrVal = Left$(StrVal, np)
        End If
        m_ServicePack = StrVal
    End If
End Sub

Friend Sub LetSuiteMask(ByVal Value As Long)
    m_wSuiteMask = Value
End Sub

Friend Sub LetPlatFormID(ByVal Value As Long)
    'm_dwPlatFormID = Value
    m_Platform = Value
End Sub

Private Function BitToStr() As String
    BitToStr = IIf(IsWow64, "64", "32") & "Bit"
End Function

Public Function IsWow64() As Boolean

    Dim lngIsWow64 As Long

    If APIFunctionExists("kernel32", "IsWow64Process") Then
        If (IsWow64Process(GetCurrentProcess, lngIsWow64) <> 0) Then
            IsWow64 = CBool(lngIsWow64)
        Else

            ' // handle error
        End If
    End If

End Function

Public Function APIFunctionExists(ByVal DllName As String, ByVal FncName As String) As Boolean

    Dim hMod As Long: hMod = GetModuleHandleW(StrPtr(DllName))
    APIFunctionExists = (GetProcAddress(hMod, FncName) <> 0)

End Function






