VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Compressor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Header Compressapi.h on Windows 8 and Windows Server 2012
'COMPRESS_ALGORITHM
Public Enum LongPtr
    [_Value]
End Enum
Public Enum ECompressAlgo
'#define COMPRESS_ALGORITHM_INVALID 0
'#define COMPRESS_ALGORITHM_NULL 1
    RAW = &H20000000
    MSZIP = 2
    XPRESS = 3
    XPRESS_HUFF = 4
    LZMS = 5
End Enum
'#define COMPRESS_RAW (1 << 29)

'COMPRESS_INFORMATION_CLASS
Public Enum ECompressInfoClass
    INVALID = 0
    BLOCK_SIZE = 1
    Level = 2
End Enum

'Public Declare PtrSafe Function GetLastError Lib "kernel32.dll" () As Long
'Visual Basic:  Applications should call err.LastDllError instead of GetLastError.

'typedef PVOID (__cdecl *PFN_COMPRESS_ALLOCATE) (PVOID UserContext, SIZE_T Size);
'typedef VOID (__cdecl *PFN_COMPRESS_FREE) (PVOID UserContext, PVOID Memory);

'Type COMPRESS_ALLOCATION_ROUTINES
'Public Type CompressAllocRoutines
'    pfnAllocMem As Long ' Callback that allocates memory.
'    pfnFreeMem  As Long ' Callback that deallocates memory.
'    PVOID_UserContext As Long ' A pointer to context information for the allocation or deallocation routine defined by the user.
''} COMPRESS_ALLOCATION_ROUTINES, *PCOMPRESS_ALLOCATION_ROUTINES;
'End Type

'BOOL CreateCompressor(
'  _In_     DWORD                         Algorithm,
'  _In_opt_ PCOMPRESS_ALLOCATION_ROUTINES AllocationRoutines,
'  _Out_    PCOMPRESSOR_HANDLE            CompressorHandle
');
Private Declare Function CreateCompressor Lib "Cabinet" (ByVal Algorithm As ECompressAlgo, ByVal pAllocationRoutines As LongPtr, ByRef out_CompressorHandle As LongPtr) As Long

'If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error information, call GetLastError.

'BOOL CreateDecompressor(
'  _In_     DWORD                         Algorithm,
'  _In_opt_ PCOMPRESS_ALLOCATION_ROUTINES AllocationRoutines,
'  _Out_    PDECOMPRESSOR_HANDLE          DecompressorHandle
');
Private Declare Function CreateDecompressor Lib "Cabinet" (ByVal Algorithm As ECompressAlgo, ByVal AllocationRoutines As LongPtr, ByRef out_DecompressorHandle As LongPtr) As Long

'BOOL CloseCompressor( _In_ COMPRESSOR_HANDLE CompressorHandle);
Private Declare Function CloseCompressor Lib "Cabinet" (ByVal CompressorHandle As LongPtr) As Long

'BOOL CloseCompressor( _In_ COMPRESSOR_HANDLE CompressorHandle);
Private Declare Function CloseDecompressor Lib "Cabinet" (ByVal DecompressorHandle As LongPtr) As Long


'BOOL Compress(
'  _In_  COMPRESSOR_HANDLE CompressorHandle,
'  _In_  PVOID             UncompressedData,
'  _In_  SIZE_T            UncompressedDataSize,
'  _Out_ PVOID             CompressedBuffer, 'hmmm wieso eigentlich hier out?
'  _In_  SIZE_T            CompressedBufferSize,
'  _Out_ PSIZE_T           CompressedDataSize
');
Private Declare Function Compress Lib "Cabinet" (ByVal CompressorHandle As LongPtr, ByVal pUncompressedData As LongPtr, ByVal UncompressedDataSize As Long, ByVal pCompressedBuffer As LongPtr, ByVal CompressedBufferSize As Long, ByRef CompressedDataSize As Long) As Long
'If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.

'BOOL Decompress(
'  _In_  DECOMPRESSOR_HANDLE DecompressorHandle,
'  _In_  PVOID               CompressedData,
'  _In_  SIZE_T              CompressedDataSize,
'  _Out_ PVOID               UncompressedBuffer,
'  _In_  SIZE_T              UncompressedBufferSize,
'  _Out_ PSIZE_T             UncompressedDataSize
');
Private Declare Function Decompress Lib "Cabinet" (ByVal DecompressorHandle As LongPtr, ByVal pCompressedData As LongPtr, ByVal CompressedDataSize As Long, ByVal pUncompressedBuffer As LongPtr, ByVal UncompressedBufferSize As Long, ByRef UncompressedDataSize As Long) As Long

Private m_CompHnd As LongPtr
Private m_DecpHnd As LongPtr

Private Sub Class_Initialize()
    '
End Sub
Private Sub InitCompHandle()
    'If CreateCompressor(RAW Or MSZIP, 0, m_CompHnd) = 0 Then
    If CreateCompressor(MSZIP, 0, m_CompHnd) = 0 Then MsgBox "Could not create compressor-handle."
End Sub
Private Sub InitDecompHandle()
    If CreateDecompressor(RAW Or MSZIP, 0, m_DecpHnd) = 0 Then MsgBox "Could not create decompressor-handle."
End Sub

Private Sub Class_Terminate()
    If m_CompHnd <> 0 Then CloseCompressor m_CompHnd
    If m_DecpHnd <> 0 Then CloseDecompressor m_DecpHnd
End Sub

Public Function CompressFile(ByVal apfn As String, out_Compressed() As Byte) As Boolean
    'Dim Data() As Byte: Data = LoadData(apfn)
    'CompressData Data
    CompressFile = CompressData(LoadData(apfn), out_Compressed)
End Function
Private Function LoadData(pfn As String) As Byte()
Try: On Error GoTo Finally
    Dim Data() As Byte
    Dim FNr As Integer: FNr = FreeFile
    Open pfn For Binary As FNr
    ReDim Data(0 To LOF(FNr) - 1)
    Get FNr, , Data
    LoadData = Data
Finally:
    Close FNr
    If Err Then MsgBox Err.Description
End Function
Public Function CompressData(Data() As Byte, out_Compressed() As Byte) As Boolean
    If m_CompHnd = 0 Then InitCompHandle
    If m_CompHnd <> 0 Then
        Dim pData As LongPtr: pData = VarPtr(Data(0))
        Dim DataSize As Long: DataSize = UBound(Data) + 1
        Dim CompSize As Long
        
        'in 2 Stufen
        'zuerst die Größe der compressed daten abfragen
        Dim nullptr As LongPtr
        If Compress(m_CompHnd, pData, DataSize, nullptr, nullptr, CompSize) Then
            '
        End If
        'dann die Compression durchführen
        If CompSize Then
            ReDim out_Compressed(0 To CompSize - 1) ' * 2
            Dim pCompressed As LongPtr: pCompressed = VarPtr(out_Compressed(0))
            CompressData = Compress(m_CompHnd, pData, DataSize, pCompressed, CompSize, CompSize) <> 0
            'ReDim Preserve out_Compressed(0 To CompSize - 1)
        End If
    End If
End Function
'Compress hCompressorHandle, VarPtr(bBufferIn(0)), UBound(bBufferIn), 0&, 0&, VarPtr(lCompressedDataSize)

'ReDim bBufferOut(lCompressedDataSize - 1)
                
'Compress hCompressorHandle, VarPtr(bBufferIn(0)), UBound(bBufferIn), VarPtr(bBufferOut(0)), lCompressedDataSize, VarPtr(lCompressedDataSize)

Public Function DecompressFile(ByVal apfn As String, out_Decompressed() As Byte) As Boolean
    'Dim Data() As Byte: Data = LoadData(apfn)
    'CompressData Data
    DecompressFile = DecompressData(LoadData(apfn), out_Decompressed)
End Function
Public Function DecompressData(Data() As Byte, out_Decompressed() As Byte) As Boolean
    If m_DecpHnd = 0 Then InitDecompHandle
    If m_DecpHnd <> 0 Then
        Dim pData As LongPtr: pData = VarPtr(Data(0))
        Dim DataSize As Long: DataSize = UBound(Data) + 1
        Dim DecpSize As Long
        
        'in 2 Stufen
        'zuerst die Größe der compressed daten abfragen
        Dim nullptr As LongPtr
        If Decompress(m_DecpHnd, pData, DataSize, nullptr, nullptr, DecpSize) Then
            '
        End If
        'dann die Compression durchführen
        If DecpSize > 0 Then
            ReDim out_Decompressed(0 To DecpSize - 1) '* 2
            Dim pDecompressed As LongPtr: pDecompressed = VarPtr(out_Decompressed(0))
            DecompressData = Decompress(m_DecpHnd, pData, DataSize, pDecompressed, DecpSize, DecpSize) <> 0
            'ReDim Preserve out_Decompressed(0 To DecpSize - 1)
        End If
    End If
End Function

    'With m_CAR
        '.pfnAllocMem = FncPtr(AddressOf ComprAllocMem)
        '.pfnFreeMem = FncPtr(AddressOf ComprFreeMem)
        '.PVOID_UserContext = 0
    'End With
    'RAW Or
    'If CreateCompressor(MSZIP, VarPtr(m_CAR), m_CompHnd) Then

                    
'Public Function FncPtr(ByVal pfn As Long) As Long
'    FncPtr = pfn
'End Function
'Private Function ComprAllocMem(ByVal pUserContext As LongPtr, ByVal Size As Long) As Long
'    '
'End Function
'Private Sub ComprFreeMem(ByVal pUserContext As LongPtr, ByVal PMemory As LongPtr)
'    '
'End Sub
'



