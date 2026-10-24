VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZLibWAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const INVALID_HANDLE_VALUE  As Long = -1
Private Const OPEN_EXISTING         As Long = 3
Private Const FILE_SHARE_READ       As Long = &H1
Private Const FILE_SHARE_WRITE      As Long = &H2
Private Const GENERIC_WRITE         As Long = &H40000000
Private Const GENERIC_READ          As Long = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const MAX_PATH              As Long = 260 'oh c'mon!


Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

'Private Type SYSTEMTIME
'  Year As Integer
'  Month As Integer
'  DayOfWeek As Integer
'  Day As Integer
'  Hour As Integer
'  Minute As Integer
'  Second As Integer
'  Milliseconds As Integer
'End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type


Private Declare Function RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLength As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpBuffer As String, ByVal lpString As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, CreationTime As FILETIME, LastAccessTime As FILETIME, LastWriteTime As FILETIME) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function ZLibVer Lib "zlibwapi" Alias "zlibVersion" () As Long
'Private Declare Function Compress Lib "zlibwapi.dll" Alias "compress" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
'Private Declare Function Compress2 Lib "zlibwapi.dll" Alias "compress2" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal Level As Long) As Long
Private Declare Function uncompress Lib "zlibwapi" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function lCRC32 Lib "zlibwapi" Alias "crc32" (ByVal crc As Long, Buffer As Any, ByVal Length As Long) As Long

                 
Public Enum eZipError
    zeZLibNotInstalled = 1
    zeNotZipFile = 2
    zeNoOpenZipFile = 3
    zeUnsupportedCompressionMethod = 4
    zeChecksumError = 5
    zeFileNotFound = 10
    zeFileAlreadyExists = 11
    zeCantRemoveFile = 12
    zeCantCreateFolder = 13
    zeCantCreateFile = 14
End Enum

Private Type typCentralFileHeader
    CentralFileHeaderSigniature As Long
    VersionMadeBy               As Integer
    VersionNeededToExtract      As Integer
    GeneralPurposeBitFlag       As Integer
    CompressionMethod           As Integer
    LastModFileTime             As Integer
    LastModFileDate             As Integer
    CRC32                       As Long
    CompressedSize              As Long
    UnCompressedSize            As Long
    FileNameLength              As Integer
    ExtraFieldLength            As Integer
    FileCommentLength           As Integer
    DiskNumberStart             As Integer
    InternalFileAttributes      As Integer
    ExternalFileAttributes      As Long
    RelativeOffsetOfLocalHeader As Long
End Type

Private Type typCenteralDirEnd
    EndOFCentralDirSignature         As Long
    NumberOfThisDisk                 As Integer
    NumberOfDiskWithCentralDir       As Integer
    EntriesInTheCentralDirThisOnDisk As Integer
    EntriesInTheCentralDir           As Integer
    SizeOfCentralDir                 As Long
    OffSetOfCentralDir               As Long
    ZipFileCommentLength             As Integer
End Type

Private Type typLocalFileHeader
    LocalFileHeaderSignature As Long
    VersionNeededToExtract   As Integer
    GeneralPurposeBitFlag    As Integer
    CompressionMethod        As Integer
    LastModFileTime          As Integer
    LastModFileDate          As Integer
    CRC32                    As Long
    CompressedSize           As Long
    UnCompressedSize         As Long
    FileNameLength           As Integer
    ExtraFieldLength         As Integer
End Type

Private Const EndOFCentralDirSignature    As Long = &H6054B50
Private Const CentralFileHeaderSigniature As Long = &H2014B50
Private Const LocalFileHeaderSignature    As Long = &H4034B50

Private CentralFileHeader As typCentralFileHeader
Private CentralDirEnd     As typCenteralDirEnd

Private CentralDirEndPos As Long
Private m_fID            As Integer
Private m_bOpenZip       As Boolean

' Zlib Status-Codes
Private Const Z_OK As Long = 0
Private Const Z_BUF_ERROR As Long = -5

Public Event Progress(Percent As Long, Cancel As Boolean)
Public Event status(Text As String)
Public Event ZipError(Number As eZipError, Description As String)


Private Sub Class_Terminate()
    CloseZip
End Sub

Public Property Get ZLibVersion() As String
Try: On Error GoTo Catch
    ZLibVersion = PointerToString(ZLibVer)
    Exit Property
Catch:
    RaiseEvent ZipError(zeZLibNotInstalled, "Zlib is not installed")
End Property

Public Function OpenZip(ZipPath As String) As Boolean
    m_bOpenZip = False
    RaiseEvent status("Opening Zip")
    CloseZip
    If Not FileExists(ZipPath) Then
        RaiseEvent ZipError(zeFileNotFound, "The file " & ZipPath & " doesn't exist")
        Exit Function
    End If
    m_fID = FreeFile
    Open ZipPath For Binary As #m_fID
    CentralDirEndPos = GetCentralDirEndPos(m_fID)
    If CentralDirEndPos > 0 Then
        m_bOpenZip = True
        OpenZip = True
        RaiseEvent status("Zip Opened")
    Else
        RaiseEvent ZipError(zeNotZipFile, "The file " & ZipPath & " is not a Zip file")
    End If
End Function

Public Sub CloseZip()
    If m_fID <> 0 Then
        Close #m_fID
        m_fID = 0
        RaiseEvent status("Zip Closed")
    End If
    CentralDirEndPos = 0
End Sub

Public Property Get Files() As Collection
    Dim l As Long
    Dim sFilename As String
    Set Files = New Collection
    If m_bOpenZip Then
        If ReadCentralDirEnd(CentralDirEndPos) Then
            Seek #m_fID, CentralDirEnd.OffSetOfCentralDir + 1
            For l = 1 To CentralDirEnd.EntriesInTheCentralDir
                ReadCentralFileHeader sFilename
                Files.Add sFilename, sFilename
            Next l
        End If
    End If
End Property

Public Function GetFiles() As String()
    Dim l As Long
    Dim sFilename As String
    If m_bOpenZip Then
        If ReadCentralDirEnd(CentralDirEndPos) Then
            ReDim aTemp(CentralDirEnd.EntriesInTheCentralDir - 1) As String
            Seek #m_fID, CentralDirEnd.OffSetOfCentralDir + 1
            For l = 1 To CentralDirEnd.EntriesInTheCentralDir
                ReadCentralFileHeader sFilename
                aTemp(l - 1) = sFilename
            Next l
        End If
    End If
    GetFiles = aTemp
End Function

Public Function Count() As Long
    If m_bOpenZip Then
        Count = CentralDirEnd.EntriesInTheCentralDir
    Else
        Count = -1
    End If
End Function

''' <summary>
''' Dekomprimiert ein FlateDecode / Zlib-komprimiertes Byte-Array.
''' </summary>
Public Function FlateDecodeBytes(ByRef compressedData() As Byte, Optional ByVal estimatedSize As Long = 0) As Byte()
    Dim srcLen As Long: srcLen = UBound(compressedData) - LBound(compressedData) + 1
    ' Falls keine Zielgröße geschätzt wurde, nehmen wir das 5-Fache der Quelle an
    Dim destLen As Long: destLen = IIf(estimatedSize > 0, estimatedSize, srcLen * 5 + 100)
    
    'Dim resultBuffer()
    Dim status As Long
    
    Do
        ReDim resultBuffer(0 To destLen - 1) As Byte
        ' Aufruf der uncompress-Funktion aus der DLL
        status = uncompress(resultBuffer(0), destLen, compressedData(LBound(compressedData)), srcLen)
        ' Wenn der Puffer zu klein war, vergrößern und erneut versuchen
        If status = Z_BUF_ERROR Then destLen = destLen * 2 Else Exit Do
    Loop While True
    
    If status = Z_OK Then
        ' Puffer auf exakt benötigte Größe kürzen
        If destLen > 0 Then
            ReDim Preserve resultBuffer(0 To destLen - 1)
            FlateDecodeBytes = resultBuffer
        Else
            FlateDecodeBytes = New Collection ' Gibt leeres Array zurück
        End If
    Else
        Err.Raise vbObjectError + 513, "Zlibwapi", "Fehler bei der Dekomprimierung. Code: " & status
    End If
End Function
'PDF-Sonderfall "Predictor": Falls im PDF-Objekt neben /Filter /FlateDecode auch ein Eintrag namens
'/DecodeParams mit einem /Predictor (z. B. Predictor 12) steht, reicht pures uncompress nicht aus.
'Nach dem Entpacken müssen die Daten in diesem Fall noch zeilenweise nach der PNG-Filter-Spezifikation
'(TIFF/PNG-Prädiktoren) zurückgerechnet werden.


'---------------------------------------------------------------------------------------
' Procedure : ExtractToMemory
' DateTime  : 18-11-2006
' Author    : ArnoutV
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function ExtractToMemory(ByVal sExtractFile As String, bData() As Byte) As Boolean
    Dim l As Long
    Dim FilePos As Long
    Dim sFilename As String
    If Len(ZLibVersion) = 0 Then
        Exit Function
    End If
    If CentralDirEndPos = 0 Then
        RaiseEvent ZipError(zeNoOpenZipFile, "There is no Zip File Open")
        Exit Function
    End If
    sExtractFile = LCase$(sExtractFile)
    If ReadCentralDirEnd(CentralDirEndPos) Then
        Seek #m_fID, CentralDirEnd.OffSetOfCentralDir + 1
        For l = 1 To CentralDirEnd.EntriesInTheCentralDir
            ReadCentralFileHeader sFilename
            If LCase$(GetFileName(sFilename)) = sExtractFile Then
                If CentralFileHeader.UnCompressedSize > 0 Then
                    FilePos = Seek(m_fID)
                    ExtractToMemory = ExtractArray(sFilename, bData)
                End If
                Exit For
            End If
        Next l
    End If
End Function

Private Function GetFileName(Path As String) As String
    Dim l As Long: l = InStrRev(Path, "\")
    If l > 0 Then GetFileName = Right$(Path, Len(Path) - l) Else GetFileName = Path
End Function

Private Function GetFilePath(Path As String) As String
    Dim l As Long: l = InStrRev(Path, "\")
    If l > 0 Then GetFilePath = Left$(Path, l - 1)
End Function

Private Sub ReadCentralFileHeader(FileName As String)
    Dim ExtraField As String
    Dim Comment    As String
    Get #m_fID, , CentralFileHeader
    If CentralFileHeader.CentralFileHeaderSigniature = CentralFileHeaderSigniature Then
        FileName = Space(CentralFileHeader.FileNameLength)
        Get m_fID, , FileName
        FileName = Replace(FileName, "/", "\")
        ExtraField = Space(CentralFileHeader.ExtraFieldLength)
        Get m_fID, , ExtraField
        Comment = Space(CentralFileHeader.FileCommentLength)
        Get m_fID, , Comment
    End If
End Sub

Private Function ReadCentralDirEnd(Position As Long) As Boolean
    Dim ZipComment As String
    Get #m_fID, Position, CentralDirEnd
    ZipComment = Space(CentralDirEnd.ZipFileCommentLength)
    Get #m_fID, , ZipComment
    ReadCentralDirEnd = CentralDirEnd.NumberOfThisDisk = CentralDirEnd.NumberOfDiskWithCentralDir
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExtractArray
' DateTime  : 18-11-2006
' Author    : ArnoutV
' Purpose   :
'---------------------------------------------------------------------------------------
Private Function ExtractArray(Path As String, b() As Byte) As Boolean
    Dim LocalFileHeader As typLocalFileHeader
    Dim sFilename As String
    Dim sExtraField As String
    Get m_fID, CentralFileHeader.RelativeOffsetOfLocalHeader + 1, LocalFileHeader
    If LocalFileHeader.LocalFileHeaderSignature = LocalFileHeaderSignature Then
        sFilename = Space(LocalFileHeader.FileNameLength)
        Get #m_fID, , sFilename
        sExtraField = Space(LocalFileHeader.ExtraFieldLength)
        Get #m_fID, , sExtraField
        ReDim b(LocalFileHeader.CompressedSize - 1)
        Get #m_fID, , b
        If CentralFileHeader.CompressionMethod = 0 Then
            'No Compression
            ExtractArray = True
        ElseIf CentralFileHeader.CompressionMethod = 8 Then 'Deflate Method
            If UnCompressBytes(b, LocalFileHeader.CompressedSize, LocalFileHeader.UnCompressedSize, LocalFileHeader.CRC32) Then
              ExtractArray = True
            Else
              RaiseEvent ZipError(zeChecksumError, "Data checksum error in " & Path)
            End If
        Else
            RaiseEvent ZipError(zeUnsupportedCompressionMethod, "The compression Method for " & sFilename & " is unsupported")
        End If
    End If
End Function

Private Function GetCentralDirEndPos(ByVal fID As Integer) As Long
    ReDim bData(LOF(fID) - 1) As Byte
    Get fID, , bData
    Dim l As Long, m As Long
    For l = UBound(bData) - 3 To LBound(bData) Step -1
        RtlMoveMemory m, bData(l), 4
        If m = EndOFCentralDirSignature Then
            GetCentralDirEndPos = l + 1
            Exit Function
        End If
    Next
End Function

Private Function UnCompressBytes(Buffer() As Byte, CompressedSize As Long, UnCompressedSize As Long, CRC32 As Long) As Boolean
    'Zlib's Uncompress method expects the 2 byte head that the Compress method adds
    'so we put that on first. Luckily it's always the same value.
    ReDim b(UBound(Buffer) + 2) As Byte
    b(0) = 120
    b(1) = 156
    RtlMoveMemory b(2), Buffer(0), UBound(Buffer) + 1
    Dim FileSize   As Long:   FileSize = UBound(Buffer) + 3
    Dim BufferSize As Long: BufferSize = CentralFileHeader.UnCompressedSize * 1.01 + 12
    ReDim Buffer(BufferSize - 1) As Byte
    Dim r As Long: r = uncompress(Buffer(0), BufferSize, b(0), FileSize)
    ReDim Preserve Buffer(CentralFileHeader.UnCompressedSize - 1)
    Dim DecompressCRC32 As Long: DecompressCRC32 = lCRC32(0&, Buffer(0), UBound(Buffer) + 1)
    If DecompressCRC32 = CRC32 Then
        UnCompressBytes = True
    End If
End Function

Private Function PointerToString(Pointer As Long) As String
    Dim l As Long:   l = lstrlen(Pointer)
    Dim s As String: s = Space(l)
    l = lstrcpy(s, Pointer)
    If l > 0 Then
        PointerToString = s
    End If
End Function

Private Function FileExists(sSource As String) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long: hFile = FindFirstFile(sSource, WFD)
    FileExists = hFile <> INVALID_HANDLE_VALUE
    FindClose hFile
End Function
