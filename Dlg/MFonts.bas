Attribute VB_Name = "MFonts"
Option Explicit
'needs class Font
'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-enumfontfamiliesexw
Private Declare Function EnumFontFamiliesExW Lib "gdi32" (ByVal hDC As LongPtr, ByRef lpLogFont As Any, ByVal lpEnumFontProc As Any, ByVal lParam As Collection, Optional ByVal dw As Long = 0) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByRef lpString As Byte) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long

Private m_LastFont   As Font
Private m_FontsAll   As Collection
Private m_FontsFixed As Collection
Private m_FontsVaria As Collection

Public Sub init()
    Set m_FontsAll = New Collection
    Set m_FontsFixed = New Collection
    Set m_FontsVaria = New Collection
    Dim fnt As Font: Set fnt = New Font: fnt.Clear: fnt.Charset = ECharSet.csDefault
    'MNew.Font("", 18) ', , , , , , , , , , , , , EFontPitch.fpDefault)
    Dim hwnd As LongPtr, hDC As LongPtr: hDC = GetDC(hwnd)
    EnumFontFamiliesExW hDC, ByVal fnt.Ptr, AddressOf EnumFontFamExProc, m_FontsAll
    ReleaseDC hwnd, hDC
    'Dim i As Long
    For Each fnt In m_FontsAll
        If fnt.Pitch = EFontPitch.fpFixed Then
            m_FontsFixed.Add fnt, fnt.Name
        ElseIf fnt.Pitch = EFontPitch.fpVariable Then
            m_FontsVaria.Add fnt, fnt.Name
        End If
    Next
End Sub

'int CALLBACK EnumFontFamExProc( const LOGFONT *lpelfe, const TEXTMETRIC *lpntme, DWORD FontType, lParam lParam);
Public Function EnumFontFamExProc(ByVal lpElfe As LongPtr, ByVal lpntme As LongPtr, ByVal FontType As Long, ByVal Col As Collection) As Long
    Dim NewFont As Font: Set NewFont = MNew.FontPtr(lpElfe)
    Dim nam As String: nam = NewFont.Name
    If m_LastFont Is Nothing Then
        Col.Add NewFont, nam
    Else
        If Not m_LastFont.Name = nam Then
            Col.Add NewFont, nam
        End If
    End If
    Set m_LastFont = NewFont
    EnumFontFamExProc = 1
End Function

Public Property Get FontsAll() As Collection
    If m_FontsAll Is Nothing Then MFonts.init
    'Set m_FontsAll = New Collection
    If m_FontsAll.Count = 0 Then MFonts.init
    Set FontsAll = m_FontsAll
End Property

Public Property Get FontsFixed() As Collection
    Set FontsFixed = m_FontsFixed
End Property

Public Property Get FontsVariable() As Collection
    Set FontsVariable = m_FontsVaria
End Property


