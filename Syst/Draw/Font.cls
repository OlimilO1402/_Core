VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private Const LF_FontNameSIZE As Long = 128&
'Private Const DEFAULT_CHARSET As Long = 1&

'lfWeight:
'Gewichtung / Schriftschnitt
Private Const FW_DONTCARE   As Long = 0
Private Const FW_THIN       As Long = 100 'Dünn
Private Const FW_EXTRALIGHT As Long = 200 'Extralight
'Private Const FW_ULTRALIGHT As Long = 200 'Ultralight
Private Const FW_LIGHT      As Long = 300 'Light
Private Const FW_NORMAL     As Long = 400 'Normal
'Private Const FW_REGULAR    As Long = 400 'Regulär
Private Const FW_MEDIUM     As Long = 500 'Mittel
Private Const FW_SEMIBOLD   As Long = 600 'Halbfett
'Private Const FW_DEMIBOLD   As Long = 600 'mittelfett
Private Const FW_BOLD       As Long = 700 'Fett
Private Const FW_EXTRABOLD  As Long = 800 'Extrafett
'Private Const FW_ULTRABOLD  As Long = 800 'Ultrafett
Private Const FW_HEAVY      As Long = 900 'schwer
'Private Const FW_BLACK      As Long = 900 'schwarz

'lfCharSet Zeichensatz
Private Const ANSI_CHARSET        As Long = 0   'This font supports the Windows ANSI character set.
Private Const DEFAULT_CHARSET     As Long = 1   'This font supports character set value based on the system default Windows ANSI code page. For example, when the system locale is English (United States), it is set as ANSI_CHARSET.
Private Const SYMBOL_CHARSET      As Long = 2   'This font supports the Windows symbol character set.
Private Const MAC_CHARSET         As Long = 77  'This font supports character set value based on the current system Macintosh code page. This value is used primarily in legacy code and should not generally be needed since modern Macintosh computers use Unicode for encoding.
Private Const SHIFTJIS_CHARSET    As Long = 128 'This font supports the Shift-JIS (Japanese Industry Standard) character set.
Private Const HANGEUL_CHARSET     As Long = 129 'This font supports the Korean (Hangul) character set.
Private Const JOHAB_CHARSET       As Long = 130 'This font supports the Korean (Johab) character set.
Private Const GB2312_CHARSET      As Long = 134 'This font supports the simplified (PRC) Chinese character set.
Private Const CHINESEBIG5_CHARSET As Long = 136 'This font supports the traditional Chinese (Big 5) character set.
Private Const GREEK_CHARSET       As Long = 161 'This font supports the Greek character set.
Private Const TURKISH_CHARSET     As Long = 162 'This font supports the Turkish character set.
Private Const VIETNAMESE_CHARSET  As Long = 163 'This font supports the Vietnamese character set.
Private Const HEBREW_CHARSET      As Long = 177 'This font supports the Hebrew character set.
Private Const ARABIC_CHARSET      As Long = 178 'This font supports the Arabic character set.
Private Const BALTIC_CHARSET      As Long = 186 'This font supports the Baltic character set.
Private Const RUSSIAN_CHARSET     As Long = 204 'This font supports the Cyrillic character set.
Private Const THAI_CHARSET        As Long = 222 'This font supports the Thai character set.
Private Const EASTEUROPE_CHARSET  As Long = 238 'This font supports the Eastern European character set.
Private Const OEM_CHARSET         As Long = 255 'This font supports an OEM-specific character set. The OEM character set is system dependent.

'lfOutPrecision Ausgabe-Präzision
Private Const OUT_DEFAULT_PRECIS   As Long = 0 'Gibt das Standardverhalten der Schriftartzuordnung an.
Private Const OUT_STRING_PRECIS    As Long = 1 'Dieser Wert wird nicht vom Schriftartzuordnungs-Element verwendet, aber er wird zurückgegeben, wenn Rasterschriftarten aufgelistet werden.
Private Const OUT_CHARACTER_PRECIS As Long = 2 'Wird nicht verwendet.
Private Const OUT_STROKE_PRECIS    As Long = 3 'Dieser Wert wird nicht vom Schriftartzuordnungs-Objekt verwendet, er wird jedoch zurückgegeben, wenn TrueType, andere gliederungsbasierte Schriftarten und Vektorschriftarten aufgelistet werden.
Private Const OUT_TT_PRECIS        As Long = 4 'Weist die Schriftartzuordnung an, eine TrueType-Schriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
Private Const OUT_DEVICE_PRECIS    As Long = 5 'Weist die Schriftartzuordnung an, eine Geräteschriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
Private Const OUT_RASTER_PRECIS    As Long = 6 'Weist die Schriftartzuordnung an, eine Rasterschriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
Private Const OUT_TT_ONLY_PRECIS   As Long = 7 'Weist die Schriftartzuordnung an, nur aus TrueType-Schriftarten auszuwählen. Wenn keine TrueType-Schriftarten im System installiert sind, kehrt die Schriftartzuordnung zum Standardverhalten zurück.
Private Const OUT_OUTLINE_PRECIS   As Long = 8 'Dieser Wert weist die Schriftartzuordnung an, zwischen TrueType und anderen gliederungsbasierten Schriftarten auszuwählen.
Private Const OUT_PS_ONLY_PRECIS   As Long = 10 'Weist die Schriftartzuordnung an, nur aus PostScript-Schriftarten auszuwählen. Wenn keine PostScript-Schriftarten im System installiert sind, kehrt die Schriftartzuordnung zum Standardverhalten zurück.

'lfClipPrecision Clipping Präzision
Private Const CLIP_DEFAULT_PRECIS   As Long = 0    'Gibt das Standard-Abschneideverhalten an.
Private Const CLIP_CHARACTER_PRECIS As Long = 1    'Wird nicht verwendet.
Private Const CLIP_STROKE_PRECIS    As Long = 2    'Wird nicht vom Schriftartzuordnungs-Objekt verwendet, sondern wird zurückgegeben, wenn Raster-, Vektor- oder TrueType-Schriftarten aufgelistet werden. Aus Gründen der Kompatibilität wird dieser Wert immer zurückgegeben, wenn Schriftarten aufgelistet werden.
Private Const CLIP_MASK             As Long = &HF& 'Wird nicht verwendet.
Private Const CLIP_LH_ANGLES        As Long = 16   'Wenn dieser Wert verwendet wird, hängt die Drehung für alle Schriftarten davon ab, ob die Ausrichtung des Koordinatensystems links- oder rechtshändig ist. Wenn nicht verwendet, drehen Geräteschriftarten immer gegen den Uhrzeigersinn, aber die Drehung anderer Schriftarten hängt von der Ausrichtung des Koordinatensystems ab.
Private Const CLIP_TT_ALWAYS        As Long = 32   'Wird nicht verwendet.
Private Const CLIP_EMBEDDED         As Long = 128  'Sie müssen dieses Flag angeben, um eine eingebettete schreibgeschützte Schriftart zu verwenden.
'Private Const CLIP_DFA_DISABLE      As Long = ?    'Windows XP SP1: Deaktiviert die Schriftartzuordnung für die Schriftart. Beachten Sie, dass dieses Flag keine Auswirkungen auf jede Plattform nach Windows Server 2003 hat.
'Private Const CLIP_DFA_OVERRIDE     As Long = ?    'Deaktiviert die Schriftartzuordnung für die Schriftart. Dies ist identisch mit CLIP_DFA_DISABLE, kann aber in einigen Situationen Probleme haben. Das empfohlene Flag für die Verwendung ist CLIP_DFA_DISABLE.

'lfQuality
Private Const DEFAULT_QUALITY        As Long = 0 'Das Aussehen der Schriftart spielt keine Rolle.
Private Const DRAFT_QUALITY          As Long = 1 'Die Darstellung der Schriftart ist weniger wichtig als bei verwendung PROOF_QUALITY. Für GDI-Rasterschriftarten ist die Skalierung aktiviert, was bedeutet, dass mehr Schriftgrößen verfügbar sind, aber die Qualität möglicherweise niedriger ist. Fett-, Kursiv-, Unterstrichen- und Streichschriftarten werden bei Bedarf synthetisiert.
Private Const PROOF_QUALITY          As Long = 2 'Die Zeichenqualität der Schriftart ist wichtiger als die exakte Übereinstimmung der Attribute für logische Schriftarten. Für GDI-Rasterschriftarten ist die Skalierung deaktiviert, und die Schrift, die der Größe am nächsten ist, wird ausgewählt. Obwohl die gewählte Schriftgröße möglicherweise nicht genau zugeordnet wird, wenn PROOF_QUALITY verwendet wird, ist die Qualität der Schriftart hoch und es gibt keine Verzerrung der Darstellung. Fett-, Kursiv-, Unterstrichen- und Streichschriftarten werden bei Bedarf synthetisiert.
Private Const NONANTIALIASED_QUALITY As Long = 3 'Schriftart wird nie antialiased.
Private Const ANTIALIASED_QUALITY    As Long = 4 'Schriftart wird immer antialiasiert, wenn die Schriftart sie unterstützt und die Schriftgröße nicht zu klein oder zu groß ist.
'Private Const CLEARTYPE_QUALITY      As Long = ? 'Wenn festgelegt, wird Text (wenn möglich) mit der ClearType-Antialiasing-Methode gerendert. Weitere Informationen finden Sie unter Hinweise.
'Private Const NONANTIALIASED_QUALITY As Long = ? 'Schriftart wird nie antialiased.

'lfPitchAndFamily
Private Const DEFAULT_PITCH  As Long = 0
Private Const FIXED_PITCH    As Long = 1
Private Const VARIABLE_PITCH As Long = 2

Private Const FF_DONTCARE    As Long = 0  'Verwenden Sie die Standardschriftart.
Private Const FF_ROMAN       As Long = &H10 '16 'Schriftarten mit variabler Strichbreite (proportional) und serifen. Ein Beispiel ist MS Serif.
Private Const FF_SWISS       As Long = &H20 '32 'Schriftarten mit variabler Strichbreite (proportional) und ohne Serifen. Ein Beispiel ist MS Sans Serif.
Private Const FF_MODERN      As Long = &H30 '48 'Schriftarten mit konstanter Strichbreite (Monospace), mit oder ohne Serifen. Monospace-Schriftarten sind in der Regel modern. Pica, Elite und CourierNew sind Beispiele.
Private Const FF_SCRIPT      As Long = &H40 '64 'Schriftarten, die wie Handschrift aussehen. Bespiele dafür sind Script und Cursive.
Private Const FF_DECORATIVE  As Long = &H50 '80 'Schriftarten der Neuheit. Ein Beispiel ist Old English.


Public Enum EFontWeight
    fwDontCare = 0
    fwThin = 100         'Dünn
    fwExtraLight = 200   'Extralight
    'fwUltraLight = 200   'Ultralight
    fwLight = 300        'Light
    fwNormal = 400       'Normal
    'fwRegular = 400      'Regulär
    fwMedium = 500       'Mittel
    fwSemiBold = 600     'Halbfett
    'fwDemiBold = 600     'mittelfett
    fwBold = 700         'Fett
    fwExtraBold = 800    'Extrafett
    'fwUltraBold = 800    'Ultrafett
    fwHeavy = 900        'Schwer
    'fwBlack = 900        'schwarz
End Enum

Public Enum ECharSet
    csANSI = 0          'ANSI_CHARSET
    csDefault = 1       'DEFAULT_CHARSET
    csSymbol = 2        'SYMBOL_CHARSET
    csMAC = 77          'MAC_CHARSET
    csShiftJIS = 128    'SHIFTJIS_CHARSET
    csHangul = 129      'HANGUL_CHARSET
    csJohab = 130       'JOHAB_CHARSET
    csGB2312 = 134      'GB2312_CHARSET
    csChineseBig5 = 136 'CHINESEBIG5_CHARSET
    csGreek = 161       'GREEK_CHARSET
    csTurkish = 162     'TURKISH_CHARSET
    csVetnamese = 163   'VIETNAMESE_CHARSET
    csHebrew = 177      'HEBREW_CHARSET
    csArabic = 178      'ARABIC_CHARSET
    csBaltic = 186      'BALTIC_CHARSET
    csRussian = 204     'RUSSIAN_CHARSET
    csThai = 222        'THAI_CHARSET
    csEastEurope = 238  'EASTEUROPE_CHARSET
    csOEM = 255         'OEM_CHARSET
End Enum

Public Enum EOutPrecision
    opDefault = 0    'Gibt das Standardverhalten der Schriftartzuordnung an.
    opString = 1     'Dieser Wert wird nicht vom Schriftartzuordnungs-Element verwendet, aber er wird zurückgegeben, wenn Rasterschriftarten aufgelistet werden.
    'opCharacter = 2  'Wird nicht verwendet.
    opStroke = 3     'Dieser Wert wird nicht vom Schriftartzuordnungs-Objekt verwendet, er wird jedoch zurückgegeben, wenn TrueType, andere gliederungsbasierte Schriftarten und Vektorschriftarten aufgelistet werden.
    opTT = 4         'Weist die Schriftartzuordnung an, eine TrueType-Schriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
    opDevice = 5     'Weist die Schriftartzuordnung an, eine Geräteschriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
    opRaster = 6     'Weist die Schriftartzuordnung an, eine Rasterschriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
    opTTonly = 7     'Weist die Schriftartzuordnung an, nur aus TrueType-Schriftarten auszuwählen. Wenn keine TrueType-Schriftarten im System installiert sind, kehrt die Schriftartzuordnung zum Standardverhalten zurück.
    opOutline = 8    'Dieser Wert weist die Schriftartzuordnung an, zwischen TrueType und anderen gliederungsbasierten Schriftarten auszuwählen.
    opPSonly = 10    'Weist die Schriftartzuordnung an, nur aus PostScript-Schriftarten auszuwählen. Wenn keine PostScript-Schriftarten im System installiert sind, kehrt die Schriftartzuordnung zum Standardverhalten zurück.
End Enum

Public Enum EClipPrecision
    cpDefault = 0      'Gibt das Standard-Abschneideverhalten an.
    'cpCharacter = 1    'Wird nicht verwendet.
    cpStroke = 2       'Wird nicht vom Schriftartzuordnungs-Objekt verwendet, sondern wird zurückgegeben, wenn Raster-, Vektor- oder TrueType-Schriftarten aufgelistet werden. Aus Gründen der Kompatibilität wird dieser Wert immer zurückgegeben, wenn Schriftarten aufgelistet werden.
    'cpMASK = 15 '&HF&  'Wird nicht verwendet.
    cpLHangles = 16    'Wenn dieser Wert verwendet wird, hängt die Drehung für alle Schriftarten davon ab, ob die Ausrichtung des Koordinatensystems links- oder rechtshändig ist. Wenn nicht verwendet, drehen Geräteschriftarten immer gegen den Uhrzeigersinn, aber die Drehung anderer Schriftarten hängt von der Ausrichtung des Koordinatensystems ab.
    'cpTTalways = 32    'Wird nicht verwendet.
    cpEMBEDDED = 128   'Sie müssen dieses Flag angeben, um eine eingebettete schreibgeschützte Schriftart zu verwenden.
End Enum

Public Enum EFontQuality
    fqDefault = 0        'Das Aussehen der Schriftart spielt keine Rolle.
    fqDraft = 1          'Die Darstellung der Schriftart ist weniger wichtig als bei verwendung PROOF_QUALITY. Für GDI-Rasterschriftarten ist die Skalierung aktiviert, was bedeutet, dass mehr Schriftgrößen verfügbar sind, aber die Qualität möglicherweise niedriger ist. Fett-, Kursiv-, Unterstrichen- und Streichschriftarten werden bei Bedarf synthetisiert.
    fqProof = 2          'Die Zeichenqualität der Schriftart ist wichtiger als die exakte Übereinstimmung der Attribute für logische Schriftarten. Für GDI-Rasterschriftarten ist die Skalierung deaktiviert, und die Schrift, die der Größe am nächsten ist, wird ausgewählt. Obwohl die gewählte Schriftgröße möglicherweise nicht genau zugeordnet wird, wenn PROOF_QUALITY verwendet wird, ist die Qualität der Schriftart hoch und es gibt keine Verzerrung der Darstellung. Fett-, Kursiv-, Unterstrichen- und Streichschriftarten werden bei Bedarf synthetisiert.
    fqNonAntialiased = 3 'Schriftart wird nie antialiased.
    fqAntialiased = 4    'Schriftart wird immer antialiasiert, wenn die Schriftart sie unterstützt und die Schriftgröße nicht zu klein oder zu groß ist.
End Enum

Public Enum EFontPitch
    'The pitch and family of the font. The two low-order bits specify the pitch of the font and can be one of the following values.
    fpDefault = 0
    fpFixed = 1
    fpVariable = 2
End Enum

Public Enum EFontFamily
    'Bits 4 through 7 of the member specify the font family and can be one of the following values.
    ffDontCare = 0
    ffRoman = 16      'Schriftarten mit variabler Strichbreite (proportional) und serifen. Ein Beispiel ist MS Serif.
    ffSwiss = 32      'Schriftarten mit variabler Strichbreite (proportional) und ohne Serifen. Ein Beispiel ist MS Sans Serif.
    ffModern = 48     'Schriftarten mit konstanter Strichbreite (Monospace), mit oder ohne Serifen. Monospace-Schriftarten sind in der Regel modern. Pica, Elite und CourierNew sind Beispiele.
    ffScript = 64     'Schriftarten, die wie Handschrift aussehen. Bespiele dafür sind Script und Cursive.
    ffDecorative = 80 'Schriftarten der Neuheit. Ein Beispiel ist Old English.
End Enum

'Functions for the Enums are in module MFonts

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-logfontw
'typedef struct tagLOGFONTW {
'  LONG  lfHeight;
'  LONG  lfWidth;
'  LONG  lfEscapement;
'  LONG  lfOrientation;
'  LONG  lfWeight;
'  BYTE  lfItalic;
'  BYTE  lfUnderline;
'  BYTE  lfStrikeOut;
'  BYTE  lfCharSet;
'  BYTE  lfOutPrecision;
'  BYTE  lfClipPrecision;
'  BYTE  lfQuality;
'  BYTE  lfPitchAndFamily;
'  WCHAR lfFaceName[LF_FACESIZE];
'} LOGFONTW, *PLOGFONTW, *NPLOGFONTW, *LPLOGFONTW;
Private Const LF_FACESIZE As Long = 32 * 2 * 2 '=128
Private Type LOGFONT
    lfHeight         As Long  ' 4 'Die Höhe der Zeichenzelle oder des Zeichens der Schriftart in logischen Einheiten. Der Zeichenhöhenwert (auch als em-height bezeichnet) ist der Wert der Zeichenzellenhöhe abzüglich des inneren Leitwerts.
                                  'Die Schriftartzuordnung interpretiert den in lfHeight angegebenen Wert wie folgt.
                                  '> 0 Die Schriftartzuordnung transformiert diesen Wert in Geräteeinheiten und gleicht ihn mit der Zellhöhe der verfügbaren Schriftarten ab.
                                  '= 0 Der Schriftartzuordnungsmodul verwendet einen Standardwert für die Höhe, wenn er nach einer Übereinstimmung sucht.
                                  '< 0 Die Schriftartzuordnung transformiert diesen Wert in Geräteeinheiten und gleicht seinen absoluten Wert mit der Zeichenhöhe der verfügbaren Schriftarten ab.
    lfWidth          As Long  ' 4 'Die durchschnittliche Breite der Zeichen in der Schriftart in logischen Einheiten.
                                  'Wenn lfWidth 0 ist, wird das Seitenverhältnis des Geräts mit dem Digitalisierungsseitenverhältnis der verfügbaren Schriftarten abgeglichen,
                                  'um die nächstgelegene Übereinstimmung zu finden, die durch den absoluten Wert der Differenz bestimmt wird.
    
    lfEscapement     As Long  ' 4 'Der Winkel in zehntel Grad zwischen dem Henisierungsvektor und der x-Achse des Geräts. Der Escapementvektor ist parallel zur Basiszeile einer Textzeile.
    lfOrientation    As Long  ' 4 'Der Winkel zwischen der Basislinie jedes Zeichens und der X-Achse des Geräts in zehntel Grad.
    lfWeight         As Long  ' 4 'Die Gewichtung der Schriftart im Bereich 0 bis 1000. Beispielsweise ist 400 normal und 700 fett formatiert. Wenn dieser Wert null ist, wird eine Standardgewichtung verwendet.
    lfItalic         As Byte  ' 1 'Eine kursiv geschriebene Schriftart, wenn sie auf TRUE festgelegt ist.
    lfUnderline      As Byte  ' 1 'Eine unterstrichene Schriftart, wenn sie auf TRUE festgelegt ist.
    lfStrikeOut      As Byte  ' 1 'Eine Streichschriftart, wenn sie auf TRUE festgelegt ist.
    lfCharSet        As Byte  ' 1
    lfOutPrecision   As Byte  ' 1
    lfClipPrecision  As Byte  ' 1
    lfQuality        As Byte  ' 1
    lfPitchAndFamily As Byte  ' 1
    'lfFontName(0 To LF_FontNameSIZE) As Byte
    lfFaceName(0 To LF_FACESIZE) As Byte
End Type                ' Sum: 30

Private Const PtPerInch  As Long = 72 'ppi
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long 'Muldiv = nNumber * nNumerator / nDenominator
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByRef lpString As Any) As Long

'                  HFONT CreateFontW             ([in] int cHeight, [in] int cWidth, [in] int cEscapement, [in] int cOrientation, [in] int cWeight, [in] DWORD bItalic, [in] DWORD bUnderline, [in] DWORD bStrikeOut, [in] DWORD iCharSet, [in] DWORD iOutPrecision, [in] DWORD iClipPrecision, [in] DWORD iQuality, [in] DWORD iPitchAndFamily, [in] LPCWSTR pszFaceName);
Private Declare Function CreateFontW Lib "gdi32" (ByVal cHeight As Long, ByVal cWidth As Long, ByVal cEscapement As Long, ByVal cOrientation As Long, ByVal cWeight As Long, ByVal bItalic As Long, ByVal bUnderline As Long, ByVal bStrikeout As Long, ByVal iCharSet As Long, ByVal iOutPrecision As Long, ByVal iClipPrecision As Long, ByVal iQuality As Long, ByVal iPitchAndFamily As Long, ByVal pszFaceName As LongPtr) As LongPtr
Private Declare Function CreateFontIndirectW Lib "gdi32" (ByRef lpLogFont As LOGFONT) As LongPtr
'Private Declare Function CreateFontIndirectExW Lib "gdi32" (ByRef ENUMLOGFONTEXDVA As ENUMLOGFONTEXDVA) As LongPtr
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long

Private m_LogObj As LOGFONT
Private m_PtSize As Single
Private m_Color  As Long
Private m_Handle As LongPtr

Private Sub Class_Initialize()
    Me.New_ "Segoe UI", 10 'Oh No don't do this man, ah oder vielleicht doch ... ?
End Sub
Private Sub Class_Terminate():  Me.Dispose:             End Sub

Friend Sub New_(Name As String, PointSize As Single, Optional Width As Long = 0, Optional Escapement As Long = 0, Optional Orientation As Long = 0, Optional Weight As EFontWeight = EFontWeight.fwDontCare, _
                Optional Color As Long = 0, Optional bItalic As Boolean = False, Optional bUnderline As Boolean = False, Optional bStrikeout As Boolean = False, Optional Charset As ECharSet = ECharSet.csDefault, _
                Optional OutPrec As EOutPrecision = EOutPrecision.opDefault, Optional ClipPrec As EClipPrecision = EClipPrecision.cpDefault, Optional Quality As EFontQuality = EFontQuality.fqDefault, Optional Pitch As EFontPitch = EFontPitch.fpDefault, Optional Family As EFontFamily = EFontFamily.ffDontCare)
    With m_LogObj
        .lfWidth = Width
        .lfEscapement = Escapement
        .lfOrientation = Orientation
        .lfWeight = Weight
        .lfItalic = VBA.Math.Abs(bItalic)
        .lfUnderline = VBA.Math.Abs(bUnderline)
        .lfStrikeOut = VBA.Math.Abs(bStrikeout)
        .lfCharSet = Charset
        .lfOutPrecision = OutPrec
        .lfClipPrecision = ClipPrec
        .lfQuality = Quality
    End With
    Me.Name = Name
    Me.Size = PointSize
    Me.Color = Color
    Me.Pitch = Pitch
    Me.Family = Family
End Sub

Friend Sub NewSF(aFont As StdFont)
    With aFont
        Me.Name = .Name
        Me.Size = .Size
        Me.Bold = .Bold
        Me.Charset = .Charset
        Me.Italic = .Italic
        Me.Strikethrough = .Strikethrough
        Me.Underline = .Underline
        Me.Weight = .Weight
    End With
End Sub

Friend Sub NewPtr(ByVal lpLogObj As LongPtr)
    RtlMoveMemory ByVal Me.Ptr, ByVal lpLogObj, LenB(m_LogObj)
    m_PtSize = -MulDiv(m_LogObj.lfHeight, PtPerInch, GetPixelsPerInchY)
End Sub

Public Sub NewC(other As Font, Optional ByVal deep As Boolean = False)
    'RtlMoveMemory ByVal Me.Ptr, ByVal other.Ptr, LenB(m_LogObj)
    'm_PtSize = other.PointSize
    With other
        NewPtr .Ptr
        m_PtSize = .Size
        m_Color = .Color
    End With
End Sub

'Copy to MNew:
'Public Function Font(ByVal Name As String, ByVal PointSize As Single, Optional ByVal Width As Long = 0, Optional ByVal Escapement As Long = 0, Optional ByVal Orientation As Long = 0, Optional ByVal Weight As EFontWeight = 0, _
'                     Optional ByVal Color As Long = 0, Optional ByVal bItalic As Boolean = False, Optional ByVal bUnderline As Boolean = False, Optional ByVal bStrikeout As Boolean = False, Optional ByVal Charset As ECharSet = ECharSet.csDefault, _
'                     Optional ByVal OutPrec As EOutPrecision = 0, Optional ByVal ClipPrec As EClipPrecision = 0, Optional ByVal Quality As EFontQuality = 0, Optional ByVal Pitch As EFontPitch = 0, Optional ByVal Family As EFontFamily = 0) As Font
'    Set Font = New Font: Font.New_ Name, PointSize, Width, Escapement, Orientation, Weight, Color, bItalic, bUnderline, bStrikeout, Charset, OutPrec, ClipPrec, Quality, Pitch, Family
'End Function
'Public Function FontSF(aStdFont As StdFont) As Font
'    Set FontSF = New Font: FontSF.NewSF aStdFont
'End Function
'Public Function FontPtr(ByVal lpLogFont As LongPtr) As Font
'    Set FontPtr = New Font: FontPtr.NewPtr lpLogFont
'End Function
'Public Function StdFont(ByVal Name As String, ByVal Size As Single, Optional ByVal Weight As Long = 0, Optional ByVal bItalic As Boolean = False, Optional ByVal bUnderline As Boolean = False, _
'                        Optional ByVal bStrikeout As Boolean = False, Optional ByVal Charset As ECharSet = 0, Optional ByVal bBold As Boolean = False) As StdFont
'    Set StdFont = New StdFont
'    With StdFont: .Name = Name: .Size = Size: .Weight = Weight: .Italic = bItalic: .Underline = bUnderline: .Strikethrough = bStrikeout: .Charset = Charset: .Bold = bBold: End With
'End Function

Public Sub Dispose():    If m_Handle Then DeleteObject m_Handle: m_Handle = 0
End Sub
Public Function Clone(Optional ByVal deep As Boolean = False) As Font:    Set Clone = New Font: Clone.NewC Me: End Function

Public Sub SetParams(Optional Name, Optional PointSize, Optional Width, Optional Escapement, Optional Orientation, Optional Weight, Optional Color, Optional bItalic, _
                     Optional bUnderline, Optional bStrikeout, Optional Charset, Optional OutPrec, Optional ClipPrec, Optional Quality, Optional Pitch, Optional Family)
    If Not IsMissing(Name) Then Me.Name = CStr(Name)
    If Not IsMissing(PointSize) Then Me.Size = CSng(PointSize)
    If Not IsMissing(Width) Then Me.Width = CLng(Width)
    If Not IsMissing(Escapement) Then Me.Escapement = CLng(Escapement)
    If Not IsMissing(Orientation) Then Me.Orientation = CLng(Orientation)
    If Not IsMissing(Weight) Then Me.Weight = CLng(Weight)
    If Not IsMissing(Color) Then Me.Color = CLng(Color)
    If Not IsMissing(bItalic) Then Me.Italic = CBool(bItalic)
    If Not IsMissing(bUnderline) Then Me.Underline = CBool(bUnderline)
    If Not IsMissing(bStrikeout) Then Me.Strikethrough = CBool(bStrikeout)
    If Not IsMissing(Charset) Then Me.Charset = Charset
    If Not IsMissing(OutPrec) Then Me.OutPrecision = CLng(OutPrec)
    If Not IsMissing(ClipPrec) Then Me.ClipPrecision = CLng(ClipPrec)
    If Not IsMissing(Quality) Then Me.Quality = CLng(Quality)
    If Not IsMissing(Pitch) Then Me.Pitch = CLng(Pitch)
    If Not IsMissing(Family) Then Me.Family = CLng(Family)
End Sub

Public Function Create() As LongPtr:    m_Handle = CreateFontIndirectW(m_LogObj):    Create = m_Handle: End Function

Public Sub Clear()
    Me.Dispose
    Dim lf As LOGFONT: m_LogObj = lf
    m_Color = 0
    m_PtSize = 0
End Sub

'NEIN NICHT AS VARIANT!, hier nur As LongPtr, weil zur Übergabe an Callback-funktion
Public Property Get Ptr() As LongPtr:        Ptr = VarPtr(m_LogObj): End Property
Public Property Get Handle() As LongPtr:  Handle = m_Handle:         End Property

Public Function ToStr() As String
    Dim s As String: s = "Font{" & vbCrLf
    With Me
        s = s & "Name       =" & .Name & vbCrLf
        s = s & "Size       =" & .Size & vbCrLf
        s = s & "Width      =" & .Width & vbCrLf
        s = s & "Escapement =" & .Escapement & vbCrLf
        s = s & "Orientation=" & .Orientation & vbCrLf
        s = s & "Weight     =" & Me.EFontWeight_ToStr(.Weight) & vbCrLf
        s = s & "Italic     =" & .Italic & vbCrLf
        s = s & "Underline  =" & .Underline & vbCrLf
        s = s & "Strikeout  =" & .Strikethrough & vbCrLf
        s = s & "Charset    =" & Me.ECharSet_ToStr(.Charset) & vbCrLf
        s = s & "OutPrec    =" & Me.EOutPrecision_ToStr(.OutPrecision) & vbCrLf
        s = s & "ClipPrec   =" & Me.EClipPrecision_ToStr(.ClipPrecision) & vbCrLf
        s = s & "Quality    =" & Me.EFontQuality_ToStr(.Quality) & vbCrLf
        s = s & "Pitch      =" & Me.EFontPitch_ToStr(.Pitch) & vbCrLf
        s = s & "Family     =" & Me.EFontFamily_ToStr(.Family) & vbCrLf
        s = s & "Color      =" & "&H" & Hex(.Color) & vbCrLf
    End With
    ToStr = s
End Function

Public Function ToStdFont() As StdFont
    Set ToStdFont = New StdFont
    With ToStdFont
        .Bold = Me.Bold
        .Charset = Me.Charset
        .Italic = Me.Italic
        .Name = Me.Name
        .Size = Me.Size
        .Strikethrough = Me.Strikethrough
        .Underline = Me.Underline
        .Weight = Me.Weight
    End With
End Function

Public Property Get Name() As String:    Name = Left$(m_LogObj.lfFaceName, lstrlenW(m_LogObj.lfFaceName(0))):            End Property
Friend Property Let Name(ByVal Value As String)
    RtlZeroMemory m_LogObj.lfFaceName(0), LF_FACESIZE + 1
    RtlMoveMemory m_LogObj.lfFaceName(0), ByVal StrPtr(Value), LenB(Value)
End Property

Public Property Get Height() As Long:                             Height = m_LogObj.lfHeight:                            End Property
Friend Property Let Height(ByVal Value As Long):                  m_LogObj.lfHeight = Value:                             End Property

' -MulDiv(m_LogObj.lfHeight, PtPerInch, GetPixelsPerInchY)
Public Property Get Size() As Single:    Size = m_PtSize: End Property
Public Property Let Size(ByVal Value As Single):    m_PtSize = Value:    m_LogObj.lfHeight = -MulDiv(CLng(m_PtSize), Me.GetPixelsPerInchY, PtPerInch): End Property

Public Function GetPixelsPerInchX() As Long
    Dim hwnd As LongPtr, hDC As LongPtr:   hDC = GetDC(hwnd)
    GetPixelsPerInchX = GetDeviceCaps(hDC, LOGPIXELSX)
    ReleaseDC hwnd, hDC
End Function
Public Function GetPixelsPerInchY() As Long
    Dim hwnd As LongPtr, hDC As LongPtr:   hDC = GetDC(hwnd)
    GetPixelsPerInchY = GetDeviceCaps(hDC, LOGPIXELSY)
    ReleaseDC hwnd, hDC
End Function

Public Property Get Width() As Long:                              Width = m_LogObj.lfWidth:                              End Property
Public Property Let Width(ByVal Value As Long):                   m_LogObj.lfWidth = Value:                              End Property

Public Property Get Escapement() As Long:                         Escapement = m_LogObj.lfEscapement:                    End Property
Friend Property Let Escapement(ByVal Value As Long):              m_LogObj.lfEscapement = Value:                         End Property

Public Property Get Orientation() As Long:                        Orientation = m_LogObj.lfOrientation:                  End Property
Friend Property Let Orientation(ByVal Value As Long):             m_LogObj.lfOrientation = Value:                        End Property

Public Property Get Weight() As Long:                             Weight = m_LogObj.lfWeight:                            End Property
Friend Property Let Weight(ByVal Value As Long):                  m_LogObj.lfWeight = Value:                             End Property

Public Property Get Bold() As Boolean:                            Bold = m_LogObj.lfWeight = EFontWeight.fwBold:                            End Property
Friend Property Let Bold(ByVal Value As Boolean):                 m_LogObj.lfWeight = IIf(Value, EFontWeight.fwBold, EFontWeight.fwNormal): End Property
'If Value Then m_LogObj.lfWeight = EFontWeight.fwBold Else m_LogObj.lfWeight = EFontWeight.fwNormal

Public Property Get Italic() As Boolean:                          Italic = m_LogObj.lfItalic:                            End Property
Friend Property Let Italic(ByVal Value As Boolean):               m_LogObj.lfItalic = VBA.Math.Abs(Value):               End Property

Public Property Get Underline() As Boolean:                       Underline = m_LogObj.lfUnderline:                      End Property
Friend Property Let Underline(ByVal Value As Boolean):            m_LogObj.lfUnderline = VBA.Math.Abs(Value):            End Property

Public Property Get Strikethrough() As Boolean:                   Strikethrough = m_LogObj.lfStrikeOut:                  End Property
Friend Property Let Strikethrough(ByVal Value As Boolean):        m_LogObj.lfStrikeOut = VBA.Math.Abs(Value):            End Property

Public Property Get Charset() As ECharSet:                        Charset = m_LogObj.lfCharSet:                          End Property
Friend Property Let Charset(ByVal Value As ECharSet)
    
    m_LogObj.lfCharSet = Min(255, VBA.Math.Abs(Value))

End Property
Private Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Public Property Get OutPrecision() As EOutPrecision:              OutPrecision = m_LogObj.lfOutPrecision:                End Property
Friend Property Let OutPrecision(ByVal Value As EOutPrecision):   m_LogObj.lfOutPrecision = Value:                       End Property

Public Property Get ClipPrecision() As EClipPrecision:            ClipPrecision = m_LogObj.lfClipPrecision:              End Property
Friend Property Let ClipPrecision(ByVal Value As EClipPrecision): m_LogObj.lfClipPrecision = Value:                      End Property

Public Property Get Quality() As EFontQuality:                    Quality = m_LogObj.lfQuality:                          End Property
Friend Property Let Quality(ByVal Value As EFontQuality):         m_LogObj.lfQuality = Value:                            End Property

Public Property Get Pitch() As EFontPitch:                        Pitch = m_LogObj.lfPitchAndFamily And &HF:             End Property
Friend Property Let Pitch(ByVal Value As EFontPitch)
    Dim fp As EFontPitch:  fp = Value And &HF
    Dim ff As EFontFamily: ff = m_LogObj.lfPitchAndFamily And &HF0
    m_LogObj.lfPitchAndFamily = fp Or ff
End Property

Public Property Get Family() As EFontFamily:                      Family = m_LogObj.lfPitchAndFamily And &HF0:           End Property
Friend Property Let Family(ByVal Value As EFontFamily)
    Dim fp As EFontPitch:  fp = m_LogObj.lfPitchAndFamily And &HF
    Dim ff As EFontFamily: ff = Value And &HF0
    m_LogObj.lfPitchAndFamily = fp Or ff
End Property

Public Property Get Color() As Long:                              Color = m_Color:                                       End Property
Friend Property Let Color(ByVal Value As Long):                   m_Color = Value:                                       End Property

' v ' ############################## ' v '     all Enums    ' v ' ############################## ' v '
' v ' ############################## ' v '    EFontWeight   ' v ' ############################## ' v '
Public Function EFontWeight_ToStr(e As EFontWeight) As String
    Dim s As String
    Select Case e
    Case EFontWeight.fwDontCare:   s = "Don't Care"   ' = 0
    Case EFontWeight.fwThin:       s = "Thin"       ' = 100  'Dünn
    Case EFontWeight.fwExtraLight: s = "ExtraLight" ' = 200  'Extralight
    'Case EFontWeight.fwUltraLight: s = "UltraLight" ' = 200  'Ultralight
    Case EFontWeight.fwLight:      s = "Light"      ' = 300  'Light
    Case EFontWeight.fwNormal:     s = "Normal"     ' = 400  'Normal
    'Case EFontWeight.fwRegular:    s = "Regular"    ' = 400  'Regulär
    Case EFontWeight.fwMedium:     s = "Medium"     ' = 500  'Mittel
    Case EFontWeight.fwSemiBold:   s = "SemiBold"   ' = 600  'Halbfett
    'Case EFontWeight.fwDemiBold:   s = "DemiBold"   ' = 600  'mittelfett
    Case EFontWeight.fwBold:       s = "Bold"       ' = 700  'Fett
    Case EFontWeight.fwExtraBold:  s = "ExtraBold"  ' = 800  'Extrafett
    'Case EFontWeight.fwUltraBold:  s = "UltraBold"  ' = 800  'Ultrafett
    Case EFontWeight.fwHeavy:      s = "Heavy"      ' = 900  'Schwer
    'Case EFontWeight.fwBlack:      s = "Black"      ' = 900  'schwarz
    End Select
    EFontWeight_ToStr = s
End Function

Public Function EFontWeight_Parse(ByVal Value As String) As EFontWeight
    Dim e As EFontWeight
    Select Case Value
    Case "Don't Care": e = EFontWeight.fwDontCare
    Case "Thin":       e = EFontWeight.fwThin
    Case "ExtraLight": e = EFontWeight.fwExtraLight
    Case "Light":      e = EFontWeight.fwLight
    Case "Normal":     e = EFontWeight.fwNormal
    Case "Medium":     e = EFontWeight.fwMedium
    Case "SemiBold":   e = EFontWeight.fwSemiBold
    Case "Bold":       e = EFontWeight.fwBold
    Case "ExtraBold":  e = EFontWeight.fwExtraBold
    Case "Heavy":      e = EFontWeight.fwHeavy
    End Select
    EFontWeight_Parse = e
End Function

Public Sub EFontWeight_ToCBLB(CBLB As ComboBox)
    With CBLB
        .AddItem EFontWeight_ToStr(EFontWeight.fwDontCare)
        .AddItem EFontWeight_ToStr(EFontWeight.fwThin)
        .AddItem EFontWeight_ToStr(EFontWeight.fwExtraLight)
        .AddItem EFontWeight_ToStr(EFontWeight.fwLight)
        .AddItem EFontWeight_ToStr(EFontWeight.fwNormal)
        .AddItem EFontWeight_ToStr(EFontWeight.fwMedium)
        .AddItem EFontWeight_ToStr(EFontWeight.fwSemiBold)
        .AddItem EFontWeight_ToStr(EFontWeight.fwBold)
        .AddItem EFontWeight_ToStr(EFontWeight.fwExtraBold)
        .AddItem EFontWeight_ToStr(EFontWeight.fwHeavy)
    End With
End Sub
' ^ ' ############################## ' ^ '    EFontWeight   ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '     ECharSet     ' v ' ############################## ' v '
Public Function ECharSet_ToStr(ByVal e As ECharSet) As String
    Dim s As String
    Select Case e
    Case ECharSet.csANSI:        s = "ANSI"         '= 0   'ANSI_CHARSET
    Case ECharSet.csDefault:     s = "Default"      '= 1   'DEFAULT_CHARSET
    Case ECharSet.csSymbol:      s = "Symbol"       '= 2   'SYMBOL_CHARSET
    Case ECharSet.csMAC:         s = "MAC"          '= 77  'MAC_CHARSET
    Case ECharSet.csShiftJIS:    s = "ShiftJIS"     '= 128 'SHIFTJIS_CHARSET
    Case ECharSet.csHangul:      s = "Hangul"       '= 129 'HANGUL_CHARSET
    Case ECharSet.csJohab:       s = "Johab"        '= 130 'JOHAB_CHARSET
    Case ECharSet.csGB2312:      s = "GB2312"       '= 134 'GB2312_CHARSET
    Case ECharSet.csChineseBig5: s = "ChineseBig5"  '= 136 'CHINESEBIG5_CHARSET
    Case ECharSet.csGreek:       s = "Greek"        '= 161 'GREEK_CHARSET
    Case ECharSet.csTurkish:     s = "Turkish"      '= 162 'TURKISH_CHARSET
    Case ECharSet.csVetnamese:   s = "Vetnamese"    '= 163 'VIETNAMESE_CHARSET
    Case ECharSet.csHebrew:      s = "Hebrew"       '= 177 'HEBREW_CHARSET
    Case ECharSet.csArabic:      s = "Arabic"       '= 178 'ARABIC_CHARSET
    Case ECharSet.csBaltic:      s = "Baltic"       '= 186 'BALTIC_CHARSET
    Case ECharSet.csRussian:     s = "Russian"      '= 204 'RUSSIAN_CHARSET
    Case ECharSet.csThai:        s = "Thai"         '= 222 'THAI_CHARSET
    Case ECharSet.csEastEurope:  s = "EastEurope"   '= 238 'EASTEUROPE_CHARSET
    Case ECharSet.csOEM:         s = "OEM"          '= 255 'OEM_CHARSET
    End Select
    ECharSet_ToStr = s
End Function

Public Function ECharSet_Parse(ByVal Value As String) As ECharSet
    Dim e As ECharSet
    Select Case Value
    Case "ANSI":        e = ECharSet.csANSI        '= 0   'ANSI_CHARSET
    Case "Default":     e = ECharSet.csDefault     '= 1   'DEFAULT_CHARSET
    Case "Symbol":      e = ECharSet.csSymbol      '= 2   'SYMBOL_CHARSET
    Case "MAC":         e = ECharSet.csMAC         '= 77  'MAC_CHARSET
    Case "ShiftJIS":    e = ECharSet.csShiftJIS    '= 128 'SHIFTJIS_CHARSET
    Case "Hangul":      e = ECharSet.csHangul      '= 129 'HANGUL_CHARSET
    Case "Johab":       e = ECharSet.csJohab       '= 130 'JOHAB_CHARSET
    Case "GB2312":      e = ECharSet.csGB2312      '= 134 'GB2312_CHARSET
    Case "ChineseBig5": e = ECharSet.csChineseBig5 '= 136 'CHINESEBIG5_CHARSET
    Case "Greek":       e = ECharSet.csGreek       '= 161 'GREEK_CHARSET
    Case "Turkish":     e = ECharSet.csTurkish     '= 162 'TURKISH_CHARSET
    Case "Vetnamese":   e = ECharSet.csVetnamese   '= 163 'VIETNAMESE_CHARSET
    Case "Hebrew":      e = ECharSet.csHebrew      '= 177 'HEBREW_CHARSET
    Case "Arabic":      e = ECharSet.csArabic      '= 178 'ARABIC_CHARSET
    Case "Baltic":      e = ECharSet.csBaltic      '= 186 'BALTIC_CHARSET
    Case "Russian":     e = ECharSet.csRussian     '= 204 'RUSSIAN_CHARSET
    Case "Thai":        e = ECharSet.csThai        '= 222 'THAI_CHARSET
    Case "EastEurope":  e = ECharSet.csEastEurope  '= 238 'EASTEUROPE_CHARSET
    Case "OEM":         e = ECharSet.csOEM         '= 255 'OEM_CHARSET
    End Select
    ECharSet_Parse = e
End Function

Public Sub ECharSet_ToCBLB(CBLB As ComboBox)
    With CBLB
        .AddItem ECharSet_ToStr(ECharSet.csANSI)        '= 0   'ANSI_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csDefault)     '= 1   'DEFAULT_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csSymbol)      '= 2   'SYMBOL_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csMAC)         '= 77  'MAC_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csShiftJIS)    '= 128 'SHIFTJIS_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csHangul)      '= 129 'HANGUL_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csJohab)       '= 130 'JOHAB_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csGB2312)      '= 134 'GB2312_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csChineseBig5) '= 136 'CHINESEBIG5_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csGreek)       '= 161 'GREEK_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csTurkish)     '= 162 'TURKISH_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csVetnamese)   '= 163 'VIETNAMESE_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csHebrew)      '= 177 'HEBREW_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csArabic)      '= 178 'ARABIC_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csBaltic)      '= 186 'BALTIC_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csRussian)     '= 204 'RUSSIAN_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csThai)        '= 222 'THAI_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csEastEurope)  '= 238 'EASTEUROPE_CHARSET
        .AddItem ECharSet_ToStr(ECharSet.csOEM)         '= 255 'OEM_CHARSET
    End With
End Sub
' ^ ' ############################## ' ^ '     ECharSet     ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '  EOutPrecision   ' v ' ############################## ' v '

Public Function EOutPrecision_ToStr(ByVal e As EOutPrecision) As String
    Dim s As String
    Select Case e
    Case EOutPrecision.opDefault: s = "Default"         ' = 0  'Gibt das Standardverhalten der Schriftartzuordnung an.
    Case EOutPrecision.opString:  s = "String"          ' = 1  'Dieser Wert wird nicht vom Schriftartzuordnungs-Element verwendet, aber er wird zurückgegeben, wenn Rasterschriftarten aufgelistet werden.
    'Case EOutPrecision.opCharacter: s = "opCharacter"  ' = 2  'Wird nicht verwendet.
    Case EOutPrecision.opStroke:  s = "Stroke"          ' = 3  'Dieser Wert wird nicht vom Schriftartzuordnungs-Objekt verwendet, er wird jedoch zurückgegeben, wenn TrueType, andere gliederungsbasierte Schriftarten und Vektorschriftarten aufgelistet werden.
    Case EOutPrecision.opTT:      s = "TrueType"        ' = 4  'Weist die Schriftartzuordnung an, eine TrueType-Schriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
    Case EOutPrecision.opDevice:  s = "Device"          ' = 5  'Weist die Schriftartzuordnung an, eine Geräteschriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
    Case EOutPrecision.opRaster:  s = "Raster"          ' = 6  'Weist die Schriftartzuordnung an, eine Rasterschriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
    Case EOutPrecision.opTTonly:  s = "TrueType only"   ' = 7  'Weist die Schriftartzuordnung an, nur aus TrueType-Schriftarten auszuwählen. Wenn keine TrueType-Schriftarten im System installiert sind, kehrt die Schriftartzuordnung zum Standardverhalten zurück.
    Case EOutPrecision.opOutline: s = "Outline"         ' = 8  'Dieser Wert weist die Schriftartzuordnung an, zwischen TrueType und anderen gliederungsbasierten Schriftarten auszuwählen.
    Case EOutPrecision.opPSonly:  s = "PostScript only" ' = 10 'Weist die Schriftartzuordnung an, nur aus PostScript-Schriftarten auszuwählen. Wenn keine PostScript-Schriftarten im System installiert sind, kehrt die Schriftartzuordnung zum Standardverhalten zurück.
    End Select
    EOutPrecision_ToStr = s
End Function

Public Function EOutPrecision_Parse(ByVal Value As String) As EOutPrecision
    Dim e As EOutPrecision
    Select Case Value
    Case "Default":         e = EOutPrecision.opDefault  '= 0   Gibt das Standardverhalten der Schriftartzuordnung an.
    Case "String":          e = EOutPrecision.opString   '= 1   Dieser Wert wird nicht vom Schriftartzuordnungs-Element verwendet, aber er wird zurückgegeben, wenn Rasterschriftarten aufgelistet werden.
    'case "Default":        e= EOutPrecision.opCharacter '= 2   Wird nicht verwendet.
    Case "Stroke":          e = EOutPrecision.opStroke   '= 3   Dieser Wert wird nicht vom Schriftartzuordnungs-Objekt verwendet, er wird jedoch zurückgegeben, wenn TrueType, andere gliederungsbasierte Schriftarten und Vektorschriftarten aufgelistet werden.
    Case "TrueType":        e = EOutPrecision.opTT       '= 4   Weist die Schriftartzuordnung an, eine TrueType-Schriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
    Case "Device":          e = EOutPrecision.opDevice   '= 5   Weist die Schriftartzuordnung an, eine Geräteschriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
    Case "Raster":          e = EOutPrecision.opRaster   '= 6   Weist die Schriftartzuordnung an, eine Rasterschriftart auszuwählen, wenn das System mehrere Schriftarten mit demselben Namen enthält.
    Case "TrueType only":   e = EOutPrecision.opTTonly   '= 7   Weist die Schriftartzuordnung an, nur aus TrueType-Schriftarten auszuwählen. Wenn keine TrueType-Schriftarten im System installiert sind, kehrt die Schriftartzuordnung zum Standardverhalten zurück.
    Case "Outline":         e = EOutPrecision.opOutline  '= 8   Dieser Wert weist die Schriftartzuordnung an, zwischen TrueType und anderen gliederungsbasierten Schriftarten auszuwählen.
    Case "PostScript only": e = EOutPrecision.opPSonly   '= 10  Weist die Schriftartzuordnung an, nur aus PostScript-Schriftarten auszuwählen. Wenn keine PostScript-Schriftarten im System installiert sind, kehrt die Schriftartzuordnung zum Standardverhalten zurück.
    End Select
    EOutPrecision_Parse = e
End Function

Public Sub EOutPrecision_ToCBLB(CBLB As ComboBox)
    With CBLB
        .AddItem EOutPrecision_ToStr(EOutPrecision.opDefault)
        .AddItem EOutPrecision_ToStr(EOutPrecision.opString)
        .AddItem EOutPrecision_ToStr(EOutPrecision.opStroke)
        .AddItem EOutPrecision_ToStr(EOutPrecision.opTT)
        .AddItem EOutPrecision_ToStr(EOutPrecision.opDevice)
        .AddItem EOutPrecision_ToStr(EOutPrecision.opRaster)
        .AddItem EOutPrecision_ToStr(EOutPrecision.opTTonly)
        .AddItem EOutPrecision_ToStr(EOutPrecision.opOutline)
        .AddItem EOutPrecision_ToStr(EOutPrecision.opPSonly)
    End With
End Sub
' ^ ' ############################## ' ^ '  EOutPrecision   ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '  EClipPrecision  ' v ' ############################## ' v '
Public Function EClipPrecision_ToStr(ByVal e As EClipPrecision) As String
    Dim s As String
    Select Case e
    Case EClipPrecision.cpDefault:  s = "Default"        '= 0    'Gibt das Standard-Abschneideverhalten an.
    'Case EOutPrecision.cpCharacter: s = "Character"     '= 1    'Wird nicht verwendet.
    Case EClipPrecision.cpStroke:    s = "Stroke"         '= 2    'Wird nicht vom Schriftartzuordnungs-Objekt verwendet, sondern wird zurückgegeben, wenn Raster-, Vektor- oder TrueType-Schriftarten aufgelistet werden. Aus Gründen der Kompatibilität wird dieser Wert immer zurückgegeben, wenn Schriftarten aufgelistet werden.
    'Case EOutPrecision.cpMASK:      s = "Mask"          '= &HF& 'Wird nicht verwendet.
    Case EClipPrecision.cpLHangles:  s = "LH-Angles"      '= 16   'Wenn dieser Wert verwendet wird, hängt die Drehung für alle Schriftarten davon ab, ob die Ausrichtung des Koordinatensystems links- oder rechtshändig ist. Wenn nicht verwendet, drehen Geräteschriftarten immer gegen den Uhrzeigersinn, aber die Drehung anderer Schriftarten hängt von der Ausrichtung des Koordinatensystems ab.
    'Case EOutPrecision.cpTTalways: s = "TrueType immer" '= 32   'Wird nicht verwendet.
    Case EClipPrecision.cpEMBEDDED:  s = "Embedded"       '= 128   'Sie müssen dieses Flag angeben, um eine eingebettete schreibgeschützte Schriftart zu verwenden.
    End Select
    EClipPrecision_ToStr = s
End Function

Public Function EClipPrecision_Parse(ByVal s As String) As EClipPrecision
    Dim e As EClipPrecision
    Select Case s
    Case "Default":   e = EClipPrecision.cpDefault
    Case "Stroke":    e = EClipPrecision.cpStroke
    Case "LH-Angles": e = EClipPrecision.cpLHangles
    Case "Embedded":  e = EClipPrecision.cpEMBEDDED
    End Select
    EClipPrecision_Parse = e
End Function

Public Sub EClipPrecision_ToCBLB(CBLB As ComboBox)
    With CBLB
        .AddItem EClipPrecision_ToStr(EClipPrecision.cpDefault)
        .AddItem EClipPrecision_ToStr(EClipPrecision.cpStroke)
        .AddItem EClipPrecision_ToStr(EClipPrecision.cpLHangles)
        .AddItem EClipPrecision_ToStr(EClipPrecision.cpEMBEDDED)
    End With
End Sub
' ^ ' ############################## ' ^ '  EClipPrecision  ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '   EFontQuality   ' v ' ############################## ' v '
Public Function EFontQuality_ToStr(ByVal e As EFontQuality) As String
    Dim s As String
    Select Case e
    Case EFontQuality.fqDefault:        s = "Default"
    Case EFontQuality.fqDraft:          s = "Draft"
    Case EFontQuality.fqProof:          s = "Proof"
    Case EFontQuality.fqNonAntialiased: s = "NonAntialiased"
    Case EFontQuality.fqAntialiased:    s = "Antialiased"
    End Select
    EFontQuality_ToStr = s
End Function

Public Function EFontQuality_Parse(ByVal s As String) As EFontQuality
    Dim e As EFontQuality
    Select Case s
    Case "Default":        e = EFontQuality.fqDefault
    Case "Draft":          e = EFontQuality.fqDraft
    Case "Proof":          e = EFontQuality.fqProof
    Case "NonAntialiased": e = EFontQuality.fqNonAntialiased
    Case "Antialiased":    e = EFontQuality.fqAntialiased
    End Select
    EFontQuality_Parse = e
End Function

Public Sub EFontQuality_ToCBLB(CBLB As ComboBox)
    With CBLB
        .AddItem EFontQuality_ToStr(EFontQuality.fqDefault)
        .AddItem EFontQuality_ToStr(EFontQuality.fqDraft)
        .AddItem EFontQuality_ToStr(EFontQuality.fqProof)
        .AddItem EFontQuality_ToStr(EFontQuality.fqNonAntialiased)
        .AddItem EFontQuality_ToStr(EFontQuality.fqAntialiased)
    End With
End Sub
' ^ ' ############################## ' ^ '   EFontQuality   ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '    EFontPitch    ' v ' ############################## ' v '
Public Function EFontPitch_ToStr(ByVal e As EFontPitch) As String
    Dim s As String
    Select Case e
    Case EFontPitch.fpDefault:  s = "Default"
    Case EFontPitch.fpFixed:    s = "Fixed"
    Case EFontPitch.fpVariable: s = "Variable"
    End Select
    EFontPitch_ToStr = s
End Function

Public Function EFontPitch_Parse(ByVal s As String) As EFontPitch
    Dim e As EFontPitch
    Select Case s
    Case "Default":  e = EFontPitch.fpDefault
    Case "Fixed":    e = EFontPitch.fpFixed
    Case "Variable": e = EFontPitch.fpVariable
    End Select
    EFontPitch_Parse = e
End Function

Public Sub EFontPitch_ToCBLB(CBLB As ComboBox)
    With CBLB
        .AddItem EFontPitch_ToStr(EFontPitch.fpDefault)
        .AddItem EFontPitch_ToStr(EFontPitch.fpFixed)
        .AddItem EFontPitch_ToStr(EFontPitch.fpVariable)
    End With
End Sub
' ^ ' ############################## ' ^ '    EFontPitch    ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '    EFontFamily   ' v ' ############################## ' v '
Public Function EFontFamily_ToStr(ByVal e As EFontFamily) As String
    Dim s As String
    Select Case e
    Case EFontFamily.ffDontCare:   s = "Don't Care"
    Case EFontFamily.ffRoman:      s = "Roman"
    Case EFontFamily.ffSwiss:      s = "Swiss"
    Case EFontFamily.ffModern:     s = "Modern"
    Case EFontFamily.ffScript:     s = "Script"
    Case EFontFamily.ffDecorative: s = "Decorative"
    End Select
    EFontFamily_ToStr = s
End Function

Public Function EFontFamily_Parse(ByVal s As String) As EFontFamily
    Dim e As EFontFamily
    Select Case s
    Case "Don't Care": e = EFontFamily.ffDontCare
    Case "Roman":      e = EFontFamily.ffRoman
    Case "Swiss":      e = EFontFamily.ffSwiss
    Case "Modern":     e = EFontFamily.ffModern
    Case "Script":     e = EFontFamily.ffScript
    Case "Decorative": e = EFontFamily.ffDecorative
    End Select
    EFontFamily_Parse = e
End Function

Public Sub EFontFamily_ToCBLB(CBLB As ComboBox)
    With CBLB
        .AddItem EFontFamily_ToStr(EFontFamily.ffDontCare)
        .AddItem EFontFamily_ToStr(EFontFamily.ffRoman)
        .AddItem EFontFamily_ToStr(EFontFamily.ffSwiss)
        .AddItem EFontFamily_ToStr(EFontFamily.ffModern)
        .AddItem EFontFamily_ToStr(EFontFamily.ffScript)
        .AddItem EFontFamily_ToStr(EFontFamily.ffDecorative)
    End With
End Sub
' ^ ' ############################## ' ^ '    EFontFamily   ' ^ ' ############################## ' ^ '

'
'So ziemlich jede Klasse braucht die folgenden 10 Funktionen:
'evtl von Objekt geerbt oder evtl über ein Interface IObject
'* Friend Sub New_(<Alle Pamameter der Klasse ByRef>)
'  keine eine oder mehrere dieser Funktionen
'  Alle Parameter der Klasse werden hier direkt beim Erzeugen des Objektes das erstenmal zugewiesen
'
'  im Modul MNew eine zugehörige:
'  * Public Function Class(<Alle Parameter der Klasse ByVal in gleicher Reihenfolge>) As Class
'
'* Public function Clone(Optional ByVal deep as boolean = False, optional ByVal PreserveHashes As Boolean = False)
'  erzeugt das neue Objekt ruft NewC auf und übergibt sich selbst
'  deep und PreserveHashes an alle inliegenden Objekte und vorallem List/CCollection weiterreichen
'* Public Sub NewC(Other as Class, Optional ByVal deep As Boolean = False, optional ByVal PreserveHashes As Boolean = False)
'  wird gebaucht in List/CCollection für Undo/Redo
'  überträgt alle Parameter von Other an dieses Objekt
'* Public Function ToStr(optional bLong As Boolean = False) As String
'  für ListBoxen, ComboBoxen und evtl zu Debugzwecken
'* Public Sub SetParams(<Alle Variablen der Klasse als optionale Variants>)
'    wird gebraucht in ModalDialog
'    im Objekt dann If NotIsMissing(Param) Then ...
'* Public Function Ptr()
'  gibt ObjPtr oder auch Zeiger auf einen internen Struct/Type-Variable zurück
'* Public Function Key() As String
'  gibt einen unique-Key / Ptr als String zurück darf sich nicht verändern.
'
'Ein ModalDialog braucht die Funktionen
'* Public sub UpdateView(Obj As Object)
'  weist das Dialog-interne Objekt von Obj zu und überträgt alle Daten in zugehörige Controls
'* Public Function UpdateData(Obj As Object) As Boolean
'  Überträgt alle Daten von allen Controls im Dialog an das Objekt durch Aufruf der Funktion SetParams
