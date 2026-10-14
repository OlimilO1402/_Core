VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public NotInheritable Class Environment ' Zeilen: 299
'          Inherits System.Object
'     Member von: System
'Zusammenfassung:
' Stellt Informationen und Veränderungsmöglichkeiten für die aktuelle
' Umgebung und Plattform zur Verfügung. Die Klasse kann nicht vererbt werden.
Option Explicit
'Private Declare Function GetVersionEx2 Lib "kernel32" Alias "GetVersionExW" (ByRef lpVersionInformation As LongPtr) As Long

'Special Thanks to ActiveVB Tipp 0061: KonRad L.M. Rudolph
'Klasse ist Bestandteil des .NET-Frameworks, mscorlib.System
Private Type ITEMID
    cb   As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As ITEMID
End Type

Private Const CSIDL_FLAG_CREATE             As Long = &H8000
Private Const CSIDL_DESKTOP                 As Long = &H0 ' Desktop
Private Const CSIDL_INTERNET                As Long = &H1
Private Const CSIDL_PROGRAMS                As Long = &H2 ' Programme (im Startmenü)
Private Const CSIDL_CONTROLS                As Long = &H3
Private Const CSIDL_PRINTERS                As Long = &H4
Private Const CSIDL_PERSONAL                As Long = &H5 ' Eigene Dateien
Private Const CSIDL_FAVORITES               As Long = &H6 ' Favoriten
Private Const CSIDL_STARTUP                 As Long = &H7 ' Autostart
Private Const CSIDL_RECENT                  As Long = &H8 ' Zuletzt verwendete Dokumente
Private Const CSIDL_SENDTO                  As Long = &H9 ' Senden An
Private Const CSIDL_BITBUCKET               As Long = &HA
Private Const CSIDL_STARTMENU               As Long = &HB ' Startmenü
Private Const CSIDL_MYDOCUMENTS             As Long = &HC
Private Const CSIDL_MYMUSIC                 As Long = &HD ' Eigene Musik
Private Const CSIDL_MYVIDEO                 As Long = &HE ' Eigene Videos
Private Const CSIDL_DESKTOPDIRECTORY        As Long = &H10
Private Const CSIDL_DRIVES                  As Long = &H11 ' Treiber
Private Const CSIDL_NETWORK                 As Long = &H12
Private Const CSIDL_NETHOOD                 As Long = &H13 ' Netzwerkumgebung
Private Const CSIDL_FONTS                   As Long = &H14 ' Schriftarten
Private Const CSIDL_TEMPLATES               As Long = &H15 ' Vorlagen
Private Const CSIDL_COMMON_STARTMENU        As Long = &H16 ' Starmenü
Private Const CSIDL_COMMON_PROGRAMS         As Long = &H17 ' Programme-Ordner im Startmenü
Private Const CSIDL_COMMON_STARTUP          As Long = &H18 ' Autostart
Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19 ' Desktop
Private Const CSIDL_APPDATA                 As Long = &H1A ' Anwendungsdaten
Private Const CSIDL_PRINTHOOD               As Long = &H1B   ' Druckerumgebung
Private Const CSIDL_LOCAL_APPDATA           As Long = &H1C ' Anwendungsdaten
Private Const CSIDL_ALTSTARTUP              As Long = &H1D
Private Const CSIDL_COMMON_ALTSTARTUP       As Long = &H1D
Private Const CSIDL_COMMON_FAVORITES        As Long = &H1F ' Favoriten
Private Const CSIDL_INTERNET_CACHE          As Long = &H20 ' Temporäre Internetdateien
Private Const CSIDL_COOKIES                 As Long = &H21 ' Cookies
Private Const CSIDL_HISTORY                 As Long = &H22 ' Verlauf
Private Const CSIDL_COMMON_APPDATA          As Long = &H23
Private Const CSIDL_WINDOWS                 As Long = &H24 ' Windows
Private Const CSIDL_SYSTEM                  As Long = &H25 ' System (bzw. System32)
Private Const CSIDL_PROGRAM_FILES           As Long = &H26 ' Programme
Private Const CSIDL_MYPICTURES              As Long = &H27 ' Eigene Bilder
Private Const CSIDL_PROFILE                 As Long = &H28 ' Profil
Private Const CSIDL_PROGRAM_FILES_COMMON    As Long = &H2B ' Gemeinsamme Dateien
Private Const CSIDL_COMMON_TEMPLATES        As Long = &H2D ' Vorlagen
Private Const CSIDL_COMMON_DOCUMENTS        As Long = &H2E
Private Const CSIDL_COMMON_ADMINTOOLS       As Long = &H2F ' Für alle Benutzer (All Users)
Private Const CSIDL_ADMINTOOLS              As Long = &H30
Private Const CSIDL_COMMON_MUSIC            As Long = &H35 ' Gemeinsame Musik
Private Const CSIDL_COMMON_PICTURES         As Long = &H36 ' Gemeinsame Bilder
Private Const CSIDL_COMMON_VIDEO            As Long = &H37 ' Gemeinsame Videos
Private Const CSIDL_CDBURN_AREA             As Long = &H3B ' CD Burning

Private Const NOERROR                       As Long = 0&

Private m_SArr()  As String
Private m_Capacity  As Long
Private m_Count     As Long
Private m_ChunkSize As Long 'immer 10 mehr, wenn mcount > mcapacity
'Public Enum Environment_SpecialFolder

Public Enum SpecialFolder
    Desktop = &H0                '0
    Programs = &H2               '2
    Personal = &H5               '5
    Favorites = &H6              '6
    Startup = &H7                '7
    Recent = &H8                 '8
    SendTo = &H9                 '9
    StartMenu = &HB              '11
    MyMusic = &HD                '13
    DesktopDirectory = &H10      '16
    MyComputer = &H11            '17
    Templates = &H15             '21
    ApplicationData = &H1A       '26
    LocalApplicationData = &H1C  '28
    InternetCache = &H20         '32
    Cookies = &H21               '33
    History = &H22               '34
    CommonApplicationData = &H23 '35
    System = &H25                '37
    ProgramFiles = &H26          '38
    MyPictures = &H27            '39
    CommonProgramFiles = &H2B    '43
End Enum

Private m_OS As OperatingSystem

Private Declare Function GetSystemDirectoryW Lib "kernel32" (ByVal lpBuffer As LongPtr, ByVal nSize As Long) As Long
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Achtung für NT4: die Funktion SHGetSpecialFolderPathA geht nicht unter NT4, dann nur mit IE4SP2 und auch nur damit
'Private Declare Function SHGetSpecialFolderPathA Lib "shell32.dll" (ByVal hWnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long

Private Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidl As Long, ByVal pszPath As LongPtr) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As LongPtr, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByRef pv As Any)
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As LongPtr) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

Private Declare Function GetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpBuffer As LongPtr, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long

Private Sub Class_Initialize()
    m_ChunkSize = 10:    m_Count = 0
    GainCommandLine
End Sub

Private Sub Class_Terminate()
    Clear
End Sub

Public Property Get CommandLine() As String
    CommandLine = Command$
End Property

Public Function GetCommandLineArgs() As String()
    Dim mSArr As Variant
    mSArr = m_SArr
    GetCommandLineArgs = mSArr
End Function

Public Function GetCommandLineArgsi(Index As Long) As String
    GetCommandLineArgsi = m_SArr(Index)
End Function

Private Sub Clear()
    'Erase mStrArr
End Sub

Private Sub InitStrArr()
    m_Capacity = m_ChunkSize
    ReDim m_SArr(0 To m_Capacity - 1)
End Sub

Private Sub AddStrCmd(ByVal StrVal As String)
    If m_Count = 0 Then InitStrArr
    m_Count = m_Count + 1
    If m_Count >= m_Capacity Then
        m_Capacity = m_Capacity + m_ChunkSize
        ReDim Preserve m_SArr(0 To m_Capacity - 1)
    End If
    m_SArr(m_Count - 1) = StrVal
End Sub

Private Sub GainCommandLine()
    Dim CmdString As String
    Dim ch As String, bolStr As Boolean
    Dim str As String
    Dim i As Integer
    CmdString = Command$ '& Space(1)
    'MsgBox CmdString
    'Split ist zwar eine Coole funktion, ist hier allerdings leider nur die
    'halbe miete, denn wenn ein arg in "" steht, und einen Pfad mit Spaces
    'hat dann funzt das nicht.
    'workTb = Split(CmdString, Space(1))
    If Len(CmdString) > 0 Then
        If Len(CmdString) > 3 Then
            If Left$(CmdString, 1) = """" Then
              '
            Else
                If UCase(Right$(CmdString, 3)) = "WAP" Then
                    CmdString = """" & CmdString & """"
                End If
            End If
        End If
    End If
    'Ja dummerweise ist nicht jeder Pfad automatisch in " eingepackt
    'selbst dann nicht, wenn der Pfad Leerzeichen enthält
    'also wie unterscheidet man das?
    For i = 1 To Len(CmdString)
        ch = Mid$(CmdString, i, 1)
        If Not bolStr Then
            If ch = """" Then
                bolStr = Not bolStr
            Else
                If Not ch = Space(1) Then
                    str = str & ch
                Else
                    If Not str = vbNullString Then Call AddStrCmd(str)
                    str = vbNullString
                End If
            End If
        Else 'habe String erkannt
            If ch = """" Then
                bolStr = Not bolStr
                Call AddStrCmd("""" & str & """")
                str = vbNullString
            Else
                str = str & ch
            End If
        End If
    Next
    'zum Schluß noch auf die tatsächliche Anzahl kürzen
    If m_Count > 0 Then
        ReDim Preserve m_SArr(0 To m_Count - 1)
    End If
End Sub
'Public Function GetEnumerator() As IUnknown
'  GetEnumerator = mStrArr
'End Function

'Public Shared Function GetFolderPath(ByVal folder As System.Environment.SpecialFolder) As String
Public Function GetFolderPath(ByVal folder As SpecialFolder) As String
    Dim StrDir As String
    StrDir = String(512, vbNullChar) 'Chr(0))
    If folder = System Then
          GetFolderPath = Left$(StrDir, GetSystemDirectoryW(StrPtr(StrDir), Len(StrDir)))
        '  ElseIf folder = temp Then
        '    'Left$(strTmpDir, GetTempPath(Len(strTmpDir), strTmpDir))
        '    GetFolderPath = Left$(StrDir, GetTempPath(Len(StrDir), StrDir))
    Else
        Dim L   As Long
        Dim rv  As Long
        Dim Num As Long: Num = folder
        Dim idl As ITEMIDLIST
        'RV = SHGetSpecialFolderPathA(App.hInstance, StrDir, folder, CLng(False))
        'RV = lstrlenA(StrDir)
        'GetFolderPath = Left$(StrDir, RV)
        rv = SHGetSpecialFolderLocation(App.hInstance, Num, idl)
        If rv = NOERROR Then
            rv = SHGetPathFromIDListW(ByVal idl.mkid.cb, ByVal StrPtr(StrDir))
            If rv Then
                Call CoTaskMemFree(rv)
                'rv = lstrlenA(StrDir)
                'GetFolderPath = Left$(StrDir, rv)
                L = lstrlenW(StrPtr(StrDir))
                If L > 0 Then GetFolderPath = Left$(StrDir, L)
            End If
        End If
    End If
End Function

'Private Function GetPath(Num As Long) As String
'    Dim Result As Long
'    Dim Buff As String
'    Dim idl As ITEMIDLIST
'
'    Result = SHGetSpecialFolderLocation(Me.hWnd, Num, idl)
'    If Result = NOERROR Then
'        Buff = Space$(512)
'        Result = SHGetPathFromIDList(ByVal idl.mkid.cb, ByVal Buff)
'        If Result Then
'            GetPath = Left(Buff, InStr(Buff, Chr(0)) - 1)
'        End If
'    End If
'End Function

'Public Shared ReadOnly Property OSVersion() As System.OperatingSystem
Public Property Get OsVersion() As OperatingSystem
    Set OsVersion = MOSVer.OsVersion
'    Dim OsV1 As OSVERSIONINFO
'    Dim OsVX As OSVERSIONINFOEX
'    Dim RetVal As Long
'Try: On Error GoTo Catch
'    If m_OS Is Nothing Then
'
'#If DefDebugCalc Then
'        MsgBox "m_OS is nothing"
'#End If
'
'        OsV1.dwOSVersionInfoSize = Len(OsV1) '148
'
'#If DefDebugCalc Then
'        MsgBox "OsV1.dwOSVersionInfoSize = Len(OsV1) = " & CStr(Len(OsV1))
'#End If
'
'        RetVal = GetVersionEx1(OsV1)
'
'#If DefDebugCalc Then
'        MsgBox "RetVal = " & CStr(RetVal)
'#End If
'
'        If RetVal = 1 Then
'
'#If DefDebugCalc Then
'            MsgBox "OsV1.dwPlatformId = " & CStr(OsV1.dwPlatformId)
'#End If
'
'            If OsV1.dwPlatformId < 2 Then
'                Set m_OS = MNew.OperatingSystem(OsV1.dwPlatformId, MNew.Version(OsV1.dwMajorVersion, OsV1.dwMinorVersion, OsV1.dwBuildNumber))
'            Else
'
'#If DefDebugCalc Then
'                MsgBox "OsV1.dwMajorVersion = " & CStr(OsV1.dwMajorVersion)
'                MsgBox "OsV1.dwMinorVersion = " & CStr(OsV1.dwMinorVersion)
'#End If
'
'                If OsV1.dwMajorVersion >= 5 And OsV1.dwMinorVersion >= 1 Then
'
'#If DefDebugCalc Then
'                    MsgBox "OsVX.dwOSVersionInfoSize = Len(OsVX) = " & CStr(Len(OsVX))
'#End If
'
'                    OsVX.dwOSVersionInfoSize = Len(OsVX)
'
'                    RetVal = GetVersionEx2(OsVX)
'
'#If DefDebugCalc Then
'                    MsgBox "RetVal = " & CStr(RetVal)
'#End If
'
'                    If RetVal = 1 Then
'
'#If DefDebugCalc Then
'                        MsgBox "OsVX.dwMajorVersion = " & CStr(OsVX.dwMajorVersion)
'                        MsgBox "OsVX.dwMinorVersion = " & CStr(OsVX.dwMinorVersion)
'#End If
'
'                        Set m_OS = MNew.OperatingSystem(OsVX.dwPlatformId, MNew.Version(OsVX.dwMajorVersion, OsVX.dwMinorVersion, OsVX.dwBuildNumber, 131072))
'
'#If DefDebugCalc Then
'                        MsgBox "OsVX.szCSDVersion = " & CStr(OsVX.szCSDVersion)
'#End If
'
'                        Call m_OS.LetServicePack(OsVX.szCSDVersion)
'
'#If DefDebugCalc Then
'                        MsgBox "OsVX.wSuiteMask = " & CStr(OsVX.wSuiteMask)
'#End If
'
'                        Call m_OS.LetSuiteMask(OsVX.wSuiteMask)
'                    Else
'
'                        Set m_OS = MNew.OperatingSystem(OsV1.dwPlatformId, MNew.Version(OsV1.dwMajorVersion, OsV1.dwMinorVersion, OsV1.dwBuildNumber))
'
'                    End If
'                Else
'                    Set m_OS = MNew.OperatingSystem(OsV1.dwPlatformId, MNew.Version(OsV1.dwMajorVersion, OsV1.dwMinorVersion, OsV1.dwBuildNumber))
'                End If
'            End If
'        Else
'            'kleiner wirds nicht sein
'            Set m_OS = MNew.OperatingSystem(0, MNew.Version(3, 1, 0))
'        End If
'    'Else
'    '    Set mOS = MNew.OperatingSystem(0, MNew.Version(3, 1, 0))
'    End If
'    Set OsVersion = m_OS
'    Exit Sub
'Catch:
'    'MsgBox Err.Number & " " & Err.Description
'    Call ErrHandler("Get_OsVersion")
End Property

Public Property Get Variable(aKey As String) As String
    Const bufflen As Long = 2048
    Dim StrBuff As String * bufflen
    Dim hr As Long
    hr = GetEnvironmentVariableW(StrPtr(aKey), StrPtr(StrBuff), bufflen)
    Variable = Left$(StrBuff, hr)
End Property

Public Property Let Variable(aKey As String, RHS As String)
    Dim hr As Long
    hr = SetEnvironmentVariableW(StrPtr(aKey), StrPtr(RHS))
End Property

Public Property Get PATHVariable() As String
    PATHVariable = Variable("PATH")
End Property
Public Property Let PATHVariable(RHS As String)
    Variable("PATH") = RHS
End Property

Public Sub AddPATHVariable(aAddPath As String)
    'zuerst die Pfadvariable abrufen
    Dim p As String:   p = PATHVariable
    If Len(p) = 0 Then p = PATHVariable
    'dann die neue hinzufügen imdem die gesamte Pfadvariable
    'übergeben wird
    If Len(p) > 0 Then
        'auch überprüfen ob der neue Pfad evtl schon enthalten ist!
        Dim nPos As Long: nPos = InStr(1, p, aAddPath)
        If nPos = 0 Then
            PATHVariable = p & ";" & aAddPath
        End If
    End If
End Sub

'##############################'   Locale ErrHandler   '##############################'
'Private Function ErrHandler(ByVal PrcName As String, _
'                            Optional AddInfo As String, _
'                            Optional BolLoud As Boolean = True, _
'                            Optional bErrLog As Boolean = True, _
'                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical _
'                            ) As VbMsgBoxResult
'    ErrHandler = MErr.MessError(TypeName(Me), PrcName, AddInfo, BolLoud, bErrLog, vbDecor)
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

        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)

    Else

        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function




