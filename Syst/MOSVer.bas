Attribute VB_Name = "MOSVer"
Option Explicit
Private Type TLong
    Value As Long
End Type
Private Type TOldDwVersion
    Major As Byte
    Minor As Byte 'revision
    Build As Integer
    'v3 As Byte
    'v4 As Byte
End Type
'low-order word:
'    'low-order byte: major
'    'hi -order byte: minor version (revision) number
'
'hi -oder word:
'
'If the function succeeds, the return value includes in the low-order word the major and minor version numbers of the operating system,
'and information about the operating system platform in the high-order word.
'For all platforms,
'the low-order word contains the version number of the operating system.
'The low-order byte of this word specifies the major version number, in hexadecimal notation.
'The high-order byte specifies the minor version (revision) number, in hexadecimal notation.
'The high-order bit is zero, the next 7 bits represent the build number, and the low-order byte is 5.


'#If WCHAR Then
    Private Const MAX_STR As Long = 255
    'https://learn.microsoft.com/en-us/windows/win32/api/sysinfoapi/nf-sysinfoapi-getversionexw
    'NOT_BUILD_WINDOWS_DEPRECATE DWORD GetVersion();
    Public Declare Function GetVersion Lib "kernel32" () As Long
    Public Declare Function GetVersionExW Lib "kernel32" (ByVal lpVersionInformation As LongPtr) As Long
    Private Declare Function VerSetConditionMask Lib "kernel32" (ByVal ConditionMask As Currency, ByVal TypeMask As Long, ByVal Condition As Byte) As Currency
    
'BOOL VerifyVersionInfoW(
'  [in] LPOSVERSIONINFOEXW lpVersionInformation,
'  [in] DWORD              dwTypeMask,
'  [in] DWORDLONG          dwlConditionMask
');
Private Declare Function VerifyVersionInfoW Lib "kernel32" (ByRef OSVERSIONINFOEX As OSVERSIONINFOEX, ByVal dwTypeMask As Long, ByVal dwlConditionMask As Long) As Long

'#Else
'    Private Const MAX_STR As Long = 127
'    'https://learn.microsoft.com/en-us/windows/win32/api/sysinfoapi/nf-sysinfoapi-getversionexa
'    Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByVal lpVersionInformation As LongPtr) As Long
'#End If

'https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-osversioninfoa

'typedef struct _OSVERSIONINFOA {
'  DWORD dwOSVersionInfoSize;
'  DWORD dwMajorVersion;
'  DWORD dwMinorVersion;
'  DWORD dwBuildNumber;
'  DWORD dwPlatformId;
'  CHAR  szCSDVersion[128];
'} OSVERSIONINFOA, *POSVERSIONINFOA, *LPOSVERSIONINFOA;

'typedef struct _OSVERSIONINFOW {
'  DWORD dwOSVersionInfoSize;
'  DWORD dwMajorVersion;
'  DWORD dwMinorVersion;
'  DWORD dwBuildNumber;
'  DWORD dwPlatformId;
'  WCHAR szCSDVersion[128];
'} OSVERSIONINFOW, *POSVERSIONINFOW, *LPOSVERSIONINFOW, RTL_OSVERSIONINFOW, *PRTL_OSVERSIONINFOW;

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long                         '   4
    dwMajorVersion      As Long                         '   4
    dwMinorVersion      As Long                         '   4
    dwBuildNumber       As Long                         '   4
    dwPlatformId        As Long                         '   4
    szCSDVersion(0 To MAX_STR) As Byte      ' Service Pack  ' 128
    'szCSDVersion       As String * 128 ' Service Pack  ' 128
End Type                                         ' Sum:   148

'typedef struct _OSVERSIONINFOEXA {
'  DWORD dwOSVersionInfoSize;
'  DWORD dwMajorVersion;
'  DWORD dwMinorVersion;
'  DWORD dwBuildNumber;
'  DWORD dwPlatformId;
'  CHAR  szCSDVersion[128];
'  WORD  wServicePackMajor;
'  WORD  wServicePackMinor;
'  WORD  wSuiteMask;
'  BYTE  wProductType;
'  BYTE  wReserved;
'} OSVERSIONINFOEXA, *POSVERSIONINFOEXA, *LPOSVERSIONINFOEXA;

'typedef struct _OSVERSIONINFOEXW {
'4  DWORD dwOSVersionInfoSize;
'4  DWORD dwMajorVersion;
'4  DWORD dwMinorVersion;
'4  DWORD dwBuildNumber;
'4  DWORD dwPlatformId;
'256  WCHAR szCSDVersion[128];
'2  WORD  wServicePackMajor;
'2  WORD  wServicePackMinor;
'2  WORD  wSuiteMask;
'1  BYTE  wProductType;
'1  BYTE  wReserved;
'} OSVERSIONINFOEXW, *POSVERSIONINFOEXW, *LPOSVERSIONINFOEXW, RTL_OSVERSIONINFOEXW, *PRTL_OSVERSIONINFOEXW;

Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long                         '   4
    dwMajorVersion      As Long                         '   4
    dwMinorVersion      As Long                         '   4
    dwBuildNumber       As Long                         '   4
    dwPlatformId        As Long                         '   4
    szCSDVersion(0 To MAX_STR) As Byte      ' Service Pack  'A:128; W:256
    'szCSDVersion        As String * 128 ' Service Pack  ' 256
    wServicePackMajor   As Integer                      '   2
    wServicePackMinor   As Integer                      '   2
    wSuiteMask          As Integer                      '   2
    wProductType        As Byte                         '   1
    wReserved           As Byte                         '   1
End Type                                           ' Sum: 284

'dwTypeMask:
Private Const VER_BUILDNUMBER      As Long = &H4&   'dwBuildNumber
Private Const VER_MAJORVERSION     As Long = &H2&   'dwMajorVersion
Private Const VER_MINORVERSION     As Long = &H1&   'dwMinorVersion
Private Const VER_PLATFORMID       As Long = &H8&   'dwPlatformId
Private Const VER_PRODUCT_TYPE     As Long = &H80&  'wProductType
Private Const VER_SERVICEPACKMAJOR As Long = &H20&  'wServicePackMajor
Private Const VER_SERVICEPACKMINOR As Long = &H10&  'wServicePackMinor
Private Const VER_SUITENAME        As Long = &H49&  'wSuiteMask
Private Const VER_NT_WORKSTATION   As Long = &HA00

Private Const VER_EQUAL            As Long = 1 'The current value must be equal to the specified value.
Private Const VER_GREATER          As Long = 2 'The current value must be greater than the specified value.
Private Const VER_GREATER_EQUAL    As Long = 3 'The current value must be greater than or equal to the specified value.
Private Const VER_LESS             As Long = 4 'The current value must be less than the specified value.
Private Const VER_LESS_EQUAL       As Long = 5 'The current value must be less than or equal to the specified value.
Private Const VER_AND              As Long = 6 'All product suites specified in the wSuiteMask member must be present in the current system.
Private Const VER_OR               As Long = 7 'At least one of the specified product suites must be present in the current system.

Private Const WINVER               As Integer = &HA00
Private Const WIN32_WINNT          As Integer = &HA00
Private Const WIN32_WINNT_NT4      As Integer = &H400  '// Windows NT 4.0
Private Const WIN32_WINNT_WIN2K    As Integer = &H500  '// Windows 2000
Private Const WIN32_WINNT_WINXP    As Integer = &H501  '// Windows XP
Private Const WIN32_WINNT_WS03     As Integer = &H502  '// Windows Server 2003
Private Const WIN32_WINNT_WIN6     As Integer = &H600  '// Windows Vista
Private Const WIN32_WINNT_VISTA    As Integer = &H600  '// Windows Vista
Private Const WIN32_WINNT_WS08     As Integer = &H600  '// Windows Server 2008
Private Const WIN32_WINNT_LONGHORN As Integer = &H600  '// Windows Vista
Private Const WIN32_WINNT_WIN7     As Integer = &H601  '// Windows 7
Private Const WIN32_WINNT_WIN8     As Integer = &H602  '// Windows 8
Private Const WIN32_WINNT_WINBLUE  As Integer = &H603  '// Windows 8.1
Private Const WIN32_WINNT_WINTHRESHOLD As Integer = &HA00  '// Windows 10
Private Const WIN32_WINNT_WIN10    As Integer = &HA00  '// Windows 10

Public CSDVersion As String

Private m_OS As OperatingSystem

Private Function TOldDwVersion_ToStr(this As TOldDwVersion) As String
    With this
        TOldDwVersion_ToStr = "OsVersion = " & .Major & "." & .Minor & ", Build: " & .Build & ";"
    End With
End Function

'Public Shared ReadOnly Property OSVersion() As System.OperatingSystem
Public Property Get OsVersion() As OperatingSystem
    Dim OsV1 As OSVERSIONINFO
    Dim OsVX As OSVERSIONINFOEX
    Dim RetVal As Long
Try: On Error GoTo Catch
    If m_OS Is Nothing Then
        
        Dim tl  As TLong: tl.Value = GetVersion
        Dim odv As TOldDwVersion: LSet odv = tl
        Debug.Print TOldDwVersion_ToStr(odv)
        
'#If DefDebugCalc Then
'        MsgBox "m_OS is nothing"
'#End If
        
        OsV1.dwOSVersionInfoSize = LenB(OsV1) '148
        
'#If DefDebugCalc Then
'        MsgBox "OsV1.dwOSVersionInfoSize = Len(OsV1) = " & CStr(Len(OsV1))
'#End If
        
        RetVal = GetVersionExW(VarPtr(OsV1))
        
'#If DefDebugCalc Then
'        MsgBox "RetVal = " & CStr(RetVal)
'#End If
        
        If RetVal = 1 Then
        
'#If DefDebugCalc Then
'            MsgBox "OsV1.dwPlatformId = " & CStr(OsV1.dwPlatformId)
'#End If
            
            If OsV1.dwPlatformId < 2 Then
                Set m_OS = MNew.OperatingSystem(OsV1.dwPlatformId, MNew.Version(OsV1.dwMajorVersion, OsV1.dwMinorVersion, OsV1.dwBuildNumber))
            Else
            
'#If DefDebugCalc Then
'                MsgBox "OsV1.dwMajorVersion = " & CStr(OsV1.dwMajorVersion)
'                MsgBox "OsV1.dwMinorVersion = " & CStr(OsV1.dwMinorVersion)
'#End If
                
                If OsV1.dwMajorVersion >= 5 And OsV1.dwMinorVersion >= 1 Then
                    
'#If DefDebugCalc Then
'                    MsgBox "OsVX.dwOSVersionInfoSize = Len(OsVX) = " & CStr(Len(OsVX))
'#End If
                    
                    OsVX.dwOSVersionInfoSize = LenB(OsVX)
                    
                    RetVal = GetVersionExW(VarPtr(OsVX))
                    
'#If DefDebugCalc Then
'                    MsgBox "RetVal = " & CStr(RetVal)
'#End If
                    
                    If RetVal = 1 Then
                        
'#If DefDebugCalc Then
'                        MsgBox "OsVX.dwMajorVersion = " & CStr(OsVX.dwMajorVersion)
'                        MsgBox "OsVX.dwMinorVersion = " & CStr(OsVX.dwMinorVersion)
'#End If
                        
                        Set m_OS = MNew.OperatingSystem(OsVX.dwPlatformId, MNew.Version(OsVX.dwMajorVersion, OsVX.dwMinorVersion, OsVX.dwBuildNumber))
                        
'#If DefDebugCalc Then
'                        MsgBox "OsVX.szCSDVersion = " & CStr(OsVX.szCSDVersion)
'#End If
                        
                        Call m_OS.LetServicePack(OsVX.szCSDVersion)
                        
'#If DefDebugCalc Then
'                        MsgBox "OsVX.wSuiteMask = " & CStr(OsVX.wSuiteMask)
'#End If
                        
                        Call m_OS.LetSuiteMask(OsVX.wSuiteMask)
                    Else
                        
                        Set m_OS = MNew.OperatingSystem(OsV1.dwPlatformId, MNew.Version(OsV1.dwMajorVersion, OsV1.dwMinorVersion, OsV1.dwBuildNumber))
                        
                    End If
                Else
                    Set m_OS = MNew.OperatingSystem(OsV1.dwPlatformId, MNew.Version(OsV1.dwMajorVersion, OsV1.dwMinorVersion, OsV1.dwBuildNumber))
                End If
            End If
        Else
            'kleiner wirds nicht sein
            Set m_OS = MNew.OperatingSystem(0, MNew.Version(3, 1, 0))
        End If
    'Else
    '    Set mOS = MNew.OperatingSystem(0, MNew.Version(3, 1, 0))
    End If
    Set OsVersion = m_OS
    Exit Sub
Catch:
    'MsgBox Err.Number & " " & Err.Description
    'Call ErrHandler("Get_OsVersion")
    ErrHandler "Property Get OsVersion"
End Property

'Uses OSVERSIONINFOEXW
'e.g.
'Private Sub Text1_Click()
'    Err.Clear
'    Text1.Font = "Courier New"
'    Dim s As String
'    s = s & "ist das Windows XP    oder größer?     " & Module1.IsWindowsXPOrGreater & vbCrLf
'    s = s & "ist das Windows XPSP1 oder größer?     " & Module1.IsWindowsXPSP1OrGreater & vbCrLf
'    s = s & "ist das Windows XPSP2 oder größer?     " & Module1.IsWindowsXPSP2OrGreater & vbCrLf
'    s = s & "ist das Windows XPSP3 oder größer?     " & Module1.IsWindowsXPSP3OrGreater & vbCrLf
'    s = s & "ist das Windows Vista oder größer?     " & Module1.IsWindowsVistaOrGreater & vbCrLf
'    s = s & "ist das Windows Vista SP1 oder größer? " & Module1.IsWindowsVistaSP1OrGreater & vbCrLf
'    s = s & "ist das Windows Vista SP2 oder größer? " & Module1.IsWindowsVistaSP2OrGreater & vbCrLf
'    s = s & "ist das Windows Server           ?     " & Module1.IsWindowsServer & vbCrLf
'    s = s & "ist das Windows 7     oder größer?     " & Module1.IsWindows7OrGreater & vbCrLf
'    s = s & "ist das Windows 7 SP1 oder größer?     " & Module1.IsWindows7SP1OrGreater & vbCrLf
'
'    s = s & "ist das Windows 8     oder größer?     " & Module1.IsWindows8OrGreater & vbCrLf
'    s = s & "ist das Windows 8.1   oder größer?     " & Module1.IsWindows8Point1OrGreater & vbCrLf
'
'    s = s & "ist das Windows 10    oder größer?     " & Module1.IsWindows10OrGreater & vbCrLf
'    Text1.Text = s
'    Debug.Print Err.LastDllError & " " & Err.Description
'
'End Sub
'unter windows 11 erhält man:
'ist das Windows XP    oder größer?     Wahr
'ist das Windows XPSP1 oder größer?     Wahr
'ist das Windows XPSP2 oder größer?     Wahr
'ist das Windows XPSP3 oder größer?     Wahr
'ist das Windows Vista oder größer?     Wahr
'ist das Windows Vista SP1 oder größer? Wahr
'ist das Windows Vista SP2 oder größer? Wahr
'ist das Windows Server           ?     Falsch
'ist das Windows 7     oder größer?     Wahr
'ist das Windows 7 SP1 oder größer?     Wahr
'ist das Windows 8     oder größer?     Wahr
'ist das Windows 8.1   oder größer?     Falsch
'ist das Windows 10    oder größer?     Falsch


Public Function IsWindowsVersionOrGreater(ByVal wMajorVersion As Integer, ByVal wMinorVersion As Integer, ByVal wServicePackMajor As Integer) As Boolean
    Dim osvi As OSVERSIONINFOEX 'W
    With osvi
        .dwOSVersionInfoSize = LenB(osvi)
        .dwMajorVersion = wMajorVersion
        .dwMinorVersion = wMinorVersion
        .wServicePackMajor = wServicePackMajor
    End With
    
    Dim dwlConditionMask As Currency
    dwlConditionMask = VerSetConditionMask(VerSetConditionMask(VerSetConditionMask(0, VER_MAJORVERSION, VER_GREATER_EQUAL), VER_MINORVERSION, VER_GREATER_EQUAL), VER_SERVICEPACKMAJOR, VER_GREATER_EQUAL)
    
    IsWindowsVersionOrGreater = VerifyVersionInfoW(osvi, VER_MAJORVERSION Or VER_MINORVERSION Or VER_SERVICEPACKMAJOR, dwlConditionMask) <> 0
    'CSDVersion = osvi.szCSDVersion
    'Debug.Print "szCSDVersion: " & CStr(osvi.szCSDVersion)
    'Debug.Print "Produkttyp: " & osvi.wProductType
End Function

Public Function IsWindowsXPOrGreater() As Boolean
    IsWindowsXPOrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WINXP), LOBYTE(WIN32_WINNT_WINXP), 0)
End Function

Public Function IsWindowsXPSP1OrGreater() As Boolean
    IsWindowsXPSP1OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WINXP), LOBYTE(WIN32_WINNT_WINXP), 1)
End Function
    
Public Function IsWindowsXPSP2OrGreater() As Boolean
    IsWindowsXPSP2OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WINXP), LOBYTE(WIN32_WINNT_WINXP), 2)
End Function

Public Function IsWindowsXPSP3OrGreater() As Boolean
    IsWindowsXPSP3OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WINXP), LOBYTE(WIN32_WINNT_WINXP), 3)
End Function

Public Function IsWindowsVistaOrGreater() As Boolean
    IsWindowsVistaOrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_VISTA), LOBYTE(WIN32_WINNT_VISTA), 0)
End Function

Public Function IsWindowsVistaSP1OrGreater() As Boolean
    IsWindowsVistaSP1OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_VISTA), LOBYTE(WIN32_WINNT_VISTA), 1)
End Function

Public Function IsWindowsVistaSP2OrGreater() As Boolean
    IsWindowsVistaSP2OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_VISTA), LOBYTE(WIN32_WINNT_VISTA), 2)
End Function

Public Function IsWindows7OrGreater() As Boolean
    IsWindows7OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WIN7), LOBYTE(WIN32_WINNT_WIN7), 0)
End Function

Public Function IsWindows7SP1OrGreater() As Boolean
    IsWindows7SP1OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WIN7), LOBYTE(WIN32_WINNT_WIN7), 1)
End Function

Public Function IsWindows8OrGreater() As Boolean
    IsWindows8OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WIN8), LOBYTE(WIN32_WINNT_WIN8), 0)
End Function

Public Function IsWindows8Point1OrGreater() As Boolean
    IsWindows8Point1OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WINBLUE), LOBYTE(WIN32_WINNT_WINBLUE), 0)
End Function

Public Function HIBYTE(ByVal Value As Integer) As Byte
    HIBYTE = Value \ &H100 And &HFF&
End Function
Public Function LOBYTE(ByVal Value As Integer) As Byte
    LOBYTE = Value And &HFF&
End Function

Public Function IsWindowsThresholdOrGreater() As Boolean
    IsWindowsThresholdOrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WINTHRESHOLD), LOBYTE(WIN32_WINNT_WINTHRESHOLD), 0)
End Function

Public Function IsWindows10OrGreater() As Boolean
    IsWindows10OrGreater = IsWindowsVersionOrGreater(HIBYTE(WIN32_WINNT_WINTHRESHOLD), LOBYTE(WIN32_WINNT_WINTHRESHOLD), 0)
End Function

Public Function IsWindowsServer() As Boolean
    Dim osvi As OSVERSIONINFOEX 'W
    With osvi
        .dwOSVersionInfoSize = LenB(osvi)
        .wProductType = VER_NT_WORKSTATION
    End With
    
    Dim dwlConditionMask As Currency: dwlConditionMask = VerSetConditionMask(0, VER_PRODUCT_TYPE, VER_EQUAL)

    IsWindowsServer = VerifyVersionInfoW(osvi, VER_PRODUCT_TYPE, dwlConditionMask) <> 0
End Function



'##############################'   Locale ErrHandler   '##############################'
'Private Function ErrHandler(ByVal PrcName As String, _
'                            Optional AddInfo As String, _
'                            Optional BolLoud As Boolean = True, _
'                            Optional bErrLog As Boolean = True, _
'                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical _
'                            ) As VbMsgBoxResult
'    ErrHandler = MErr.MessError("MOSVer", PrcName, AddInfo, BolLoud, bErrLog, vbDecor)
'End Function
'copy this same function to every class, form or module
'the name of the class or form will be added automatically
'in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
' v #################### v '     Local ErrHandler      ' v #################### v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                            Optional bRetry As Boolean) As VbMsgBoxResult

    If bRetry Then

        'ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)
        ErrHandler = MessErrorRetry("MOSVer", FuncName, AddInfo, WinApiError, bErrLog)

    Else

        'ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)
        ErrHandler = MessError("MOSVer", FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function





