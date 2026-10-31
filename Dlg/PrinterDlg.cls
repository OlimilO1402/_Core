VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintDialog_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#Const IsUnicode = 1

#If vBA7 = 0 Then
Private Enum LongPtr
    [_]
End Enum
#End If
' PRINTDLG_TYPE flags-Konstanten
Private Const PD_ALLPAGES                   As Long = &H0&     ' selektiert den "Alle Seiten"-OptionButton
Private Const PD_SELECTION                  As Long = &H1&     ' selektiert die "Auswahl"-Option
Private Const PD_PAGENUMS                   As Long = &H2&     ' selektiert den "Von Bis"-OptionButton
Private Const PD_NOSELECTION                As Long = &H4&     ' "Auswahl" ist deaktiviert
Private Const PD_NOPAGENUMS                 As Long = &H8&     ' "Von Bis" ist deaktiviert
Private Const PD_COLLATE                    As Long = &H10&    ' selektiert die "Sortieren"-Checkbox
Private Const PD_PRINTTOFILE                As Long = &H20&    ' selektiert die "In Datei Drucken"-Checkbox
Private Const PD_PRINTSETUP                 As Long = &H40&    ' öffnet den Drucker-Setup-Dialog
Private Const PD_NOWARNING                  As Long = &H80&    ' keine Warnungen wenn der gewählte Drucker nicht der Standard-Drucker ist
Private Const PD_RETURNDC                   As Long = &H100&   ' gibt einen Device-Kontext des Druckers zurück
Private Const PD_RETURNIC                   As Long = &H200&   ' gibt einen Informations-Kontext zu dem gewählten Drucker als hDC zurück
Private Const PD_RETURNDEFAULT              As Long = &H400&   ' lädt den Standard-Drucker, dazu müssen bei hDevMode und hDevNames eine "0" übergeben werden
Private Const PD_SHOWHELP                   As Long = &H800&   ' Zeigt den Hilfe-Button an
Private Const PD_ENABLEPRINTHOOK            As Long = &H1000&  ' Hook-Prozedur soll benutzt werden
Private Const PD_ENABLESETUPHOOK            As Long = &H2000&  ' benutzt eine Hook-Prozedur für den Setup-Dialog
Private Const PD_ENABLEPRINTTEMPLATE        As Long = &H4000&  ' benutzt Template-Namen
Private Const PD_ENABLESETUPTEMPLATE        As Long = &H8000&  ' benutzt einen Template-Namen für den Setup-Dialog
Private Const PD_ENABLEPRINTTEMPLATEHANDLE  As Long = &H10000  ' benutzt Template-Handles
Private Const PD_ENABLESETUPTEMPLATEHANDLE  As Long = &H20000  ' benutzt ein Template-Handle für den Setup-Dialog
Private Const PD_USEDEVMODECOPIESANDCOLLATE As Long = &H40000  ' berücksichtigt Geräte-Einschränkungen und gibt nur mögliche Werte der Kopien oder des Sortierens zurück
Private Const PD_DISABLEPRINTTOFILE         As Long = &H80000  ' "In Datei Drucken" deaktivieren
Private Const PD_HIDEPRINTTOFILE            As Long = &H100000 ' "In Datei Drucken"-Button verstecken
Private Const PD_NONETWORKBUTTON            As Long = &H200000 ' Netzwerk -Button verstecken
 
' DEVMODE dmFields-Konstanten
Private Const DM_ORIENTATION                As Long = &H1&       ' dmOrientation wird gefüllt/gelesen
Private Const DM_PAPERSIZE                  As Long = &H2&       ' dmPaperSize wird gefüllt/gelesen
Private Const DM_PAPERLENGTH                As Long = &H4&       ' dmPaperLenght wird gefüllt/gelesen
Private Const DM_PAPERWIDTH                 As Long = &H8&       ' dmPaperWidth wird gefüllt/gelesen
Private Const DM_SCALE                      As Long = &H10&      ' dmScale wird gefüllt/gelesen
Private Const DM_COPIES                     As Long = &H100&     ' dmCopies wird gefüllt/gelesen
Private Const DM_DEFAULTSOURCE              As Long = &H200&     ' dmDefaultSource wird gefüllt/gelesen
Private Const DM_PRINTQUALITY               As Long = &H400&     ' dmPrintQuality wird gefüllt/gelesen
Private Const DM_COLOR                      As Long = &H800&     ' dmColor wird gefüllt/gelesen
Private Const DM_DUPLEX                     As Long = &H1000&    ' dmDuplex wird gefüllt/gelesen
Private Const DM_YRESOLUTION                As Long = &H2000&    ' dmYResolution wird gefüllt/gelesen
Private Const DM_TTOPTION                   As Long = &H4000&    ' dmTTOption wird gefüllt/gelesen
Private Const DM_COLLATE                    As Long = &H8000&    ' dmCollate wird gefüllt/gelesen
Private Const DM_FORMNAME                   As Long = &H10000    ' dmFormName wird gefüllt/gelesen
Private Const DM_LOGPIXELS                  As Long = &H20000    ' dmLogPixels wird gefüllt/gelesen
Private Const DM_BITSPERPEL                 As Long = &H40000    ' dmBitsPerPixel wird gefüllt/gelesen
Private Const DM_PELSWIDTH                  As Long = &H80000    ' dmPelsWidth wird gefüllt/gelesen
Private Const DM_PELSHEIGHT                 As Long = &H100000   ' dmPelsHeight wird gefüllt/gelesen
Private Const DM_DISPLAYFLAGS               As Long = &H200000   ' dmDisplayFlags wird gefüllt/gelesen
Private Const DM_DISPLAYFREQUENCY           As Long = &H400000   ' dmDisplayFrequency wird gefüllt/gelesen
Private Const DM_ICMMETHOD                  As Long = &H800000   ' dmICMMethod wird gefüllt/gelesen
Private Const DM_ICMINTENT                  As Long = &H1000000  ' dmICMIntent wird gefüllt/gelesen
Private Const DM_MEDIATYPE                  As Long = &H2000000  ' dmMediaType wird gefüllt/gelesen
Private Const DM_DITHERTYPE                 As Long = &H4000000  ' dmDitherType wird gefüllt/gelesen
Private Const DM_PANNINGWIDTH               As Long = &H20000000 ' dmPanningWidth wird gefüllt/gelesen
Private Const DM_PANNINGHEIGHT              As Long = &H40000000 ' dmPanningHeight wird gefüllt/gelesen

' DEVMODE dmOrientation-Konstanten
Private Const DMORIENT_PORTRAIT             As Long = 1 ' Portrait  wurde gewählt (vertikal Drucken)
Private Const DMORIENT_LANDSCAPE            As Long = 2 ' Landscape wurde gewählt (horizontal Drucken)

' DEVMODE dmPaperSize-Konstanten
Private Const DMPAPER_LETTER                  As Long = 1  ' Blatt 8,5 x 11 Inch
Private Const DMPAPER_FIRST                   As Long = 1  ' Blatt 8,5 x 11 Inch
Private Const DMPAPER_LETTERSMALL             As Long = 2  ' Blatt klein (8,5 x 11 Inch)
Private Const DMPAPER_TABLOID                 As Long = 3  ' Tabloid Blatt (11 x 17 Inch)
Private Const DMPAPER_LEDGER                  As Long = 4  ' Ledger (17x11 Inch)
Private Const DMPAPER_LEGAL                   As Long = 5  ' Blatt 8,5 x 14 Inch
Private Const DMPAPER_STATEMENT               As Long = 6  ' Statement Blatt (5,5 x 8,5 Inch)
Private Const DMPAPER_EXECUTIVE               As Long = 7  ' Executive Blatt (7.25 x 10,5 Inch)
Private Const DMPAPER_A3                      As Long = 8  ' DIN A3 Blatt
Private Const DMPAPER_A4                      As Long = 9  ' DIN A4 Blatt
Private Const DMPAPER_A4SMALL                 As Long = 10 ' DIN A4 klein (210 x 297 Millimeter) Blatt
Private Const DMPAPER_A5                      As Long = 11 ' DIN A5 Blatt
Private Const DMPAPER_B4                      As Long = 12 ' B4 Blatt
Private Const DMPAPER_B5                      As Long = 13 ' B5 Blatt
Private Const DMPAPER_FOLIO                   As Long = 14 ' Folin 8,5 + 13 Inch
Private Const DMPAPER_QUARTO                  As Long = 15 ' Quarto Blatt (215 x 275 Millimeter)
Private Const DMPAPER_10X14                   As Long = 16 ' 10 x 14 Inch
Private Const DMPAPER_11X17                   As Long = 17 ' 11 x 17 Inch
Private Const DMPAPER_NOTE                    As Long = 18 ' Note Blatt (8,5 x 11 Inch)
Private Const DMPAPER_ENV_9                   As Long = 19 ' Briefumschlag 9 (3,875 x 8,875 Inch)
Private Const DMPAPER_ENV_10                  As Long = 20 ' Briefumschlag 10 (4,125 x 9,5 Inch)
Private Const DMPAPER_ENV_11                  As Long = 21 ' Briefumschlag 11 (4,5 x 10,375 Inch)
Private Const DMPAPER_ENV_12                  As Long = 22 ' Briefumschlag 12 (4,75 x 11 Inch)
Private Const DMPAPER_ENV_14                  As Long = 23 ' Briefumschlag 14 (5 x 11,5 Inch)
Private Const DMPAPER_CSHEET                  As Long = 24 ' C Blatt (17 x 22 Inch)
Private Const DMPAPER_DSHEET                  As Long = 25 ' D Blatt (22x34 Inch)
Private Const DMPAPER_ESHEET                  As Long = 26 ' E Blatt (34 x 44 Inch)
Private Const DMPAPER_ENV_DL                  As Long = 27 ' Briefumschlag 10 (4,125 x 9,5 Inch)
Private Const DMPAPER_ENV_C5                  As Long = 28 ' Briefumschlag C5 (162 x 229 Millimeter)
Private Const DMPAPER_ENV_C3                  As Long = 29 ' Briefumschlag C3 (324 x 458 Millimeter)
Private Const DMPAPER_ENV_C4                  As Long = 30 ' Briefumschlag C4 (229 x 324 Millimeter)
Private Const DMPAPER_ENV_C6                  As Long = 31 ' Briefumschlag C6 (114 x 162 Millimeter)
Private Const DMPAPER_ENV_C65                 As Long = 32 ' Briefumschlag 10 (4,125 x 9,5 Inch)
Private Const DMPAPER_ENV_B4                  As Long = 33 ' Briefumschlag B4 (250 x 353 Millimeter)
Private Const DMPAPER_ENV_B5                  As Long = 34 ' Briefumschlag B5 (176 x 250 Millimeter)
Private Const DMPAPER_ENV_B6                  As Long = 35 ' Briefumschlag B6 (176 x 125 Millimeter)
Private Const DMPAPER_ENV_ITALY               As Long = 36 ' Italienischer Briefumschlag (110 x 230 Millimeter)
Private Const DMPAPER_ENV_MONARCH             As Long = 37 ' Monarchischer Briefumschlag (3,875 x 7,5 ' Inch)
Private Const DMPAPER_ENV_PERSONAL            As Long = 38 ' Persönlicher Briefumschlag (3,625 x 6,5 ' Inch)
Private Const DMPAPER_FANFOLD_US              As Long = 39 ' US Standard Fanfold (14,875 + 11 Inch)
Private Const DMPAPER_FANFOLD_STD_GERMAN      As Long = 40 ' Deutscher Standard Fanfold (8,5 x 12 ' Inch)
Private Const DMPAPER_FANFOLD_LGL_GERMAN      As Long = 41 ' Deutscher Rechtlicher Fanfold (8,5 x 13 Inch)
Private Const DMPAPER_LAST                    As Long = 41 ' Deutscher Rechtlicher Fanfold (8,5 x 13 Inch)
Private Const DMPAPER_ISO_B4                  As Long = 42 ' ISO B4 Blatt
Private Const DMPAPER_JAPANESE_POSTCARD       As Long = 43 ' Japanische Postkarte
Private Const DMPAPER_9X11                    As Long = 44 ' 9 x 11 Inch
Private Const DMPAPER_10X11                   As Long = 45 ' 10 x 11 Inch
Private Const DMPAPER_15X11                   As Long = 46 ' 15 x 11 Inch
Private Const DMPAPER_ENV_INVITE              As Long = 47 ' Einladungs Briefumschlag

Private Const DMPAPER_LETTER_EXTRA            As Long = 50 ' Blatt Extra
Private Const DMPAPER_LEGAL_EXTRA             As Long = 51 ' Rechtlich Extra
Private Const DMPAPER_TABLOID_EXTRA           As Long = 52 ' Tabloid Extra Blatt
Private Const DMPAPER_A4_EXTRA                As Long = 53 ' DIN A4 Extra Blatt
Private Const DMPAPER_LETTER_TRANSVERSE       As Long = 54 ' Blatt querlaufend
Private Const DMPAPER_A4_TRANSVERSE           As Long = 55 ' DIN A4 querlaufend
Private Const DMPAPER_LETTER_EXTRA_TRANSVERSE As Long = 56  ' Blatt Extra querlaufend
Private Const DMPAPER_A_PLUS                  As Long = 57 ' A Plus Blatt
Private Const DMPAPER_B_PLUS                  As Long = 58 ' B Plus Blatt
Private Const DMPAPER_LETTER_PLUS             As Long = 59 ' Blatt Plus
Private Const DMPAPER_A4_PLUS                 As Long = 60 ' DIN A4 Plus Blatt
Private Const DMPAPER_A5_TRANSVERSE           As Long = 61 ' DIN A5 querlaufend
Private Const DMPAPER_B5_TRANSVERSE           As Long = 62 ' B5 Blatt querlaufend
Private Const DMPAPER_A3_EXTRA                As Long = 63 ' DIN A3 Extra Blatt
Private Const DMPAPER_A5_EXTRA                As Long = 64 ' DIN A5 Extra Blatt
Private Const DMPAPER_B5_EXTRA                As Long = 65 ' B5 Extra Blatt
Private Const DMPAPER_A2                      As Long = 66 ' DIN A2 Blatt
Private Const DMPAPER_A3_TRANSVERSE           As Long = 67 ' DIN A3 Blatt querlaufend
Private Const DMPAPER_A3_EXTRA_TRANSVERSE     As Long = 68 ' DIN A3 Extra Blatt querlaufend
Private Const DMPAPER_USER                    As Long = 256 ' benutzerdefinierte Größe
 
' DEVMODE dmDefaultSource-Konstanten
Private Const DMBIN_ONLYONE       As Long = 1 ' nur ein Blatt
Private Const DMBIN_UPPER         As Long = 1  ' Behälter Oben
Private Const DMBIN_LOWER         As Long = 2 ' Behälter Unten
Private Const DMBIN_MIDDLE        As Long = 3 ' Behälter Mitte
Private Const DMBIN_MANUAL        As Long = 4 ' Behälter mit manueller Füllung
Private Const DMBIN_ENVELOPE      As Long = 5  ' Briefumschlag Behälter
Private Const DMBIN_ENVMANUAL     As Long = 6  ' Briefumschlag Behälter mit Manueller Füllung
Private Const DMBIN_AUTO          As Long = 7 ' automatisches füllen
Private Const DMBIN_TRACTOR       As Long = 8 ' Blatt Einzug
Private Const DMBIN_SMALLFMT      As Long = 9  ' kleines Format laden
Private Const DMBIN_LARGEFMT      As Long = 10 ' großes Format laden
Private Const DMBIN_LARGECAPACITY As Long = 11  ' große Kapazität
Private Const DMBIN_CASSETTE      As Long = 14 ' Blatt Kassette
Private Const DMBIN_FORMSOURCE    As Long = 15  ' Form Papier Quelle

' DEVMODE dmPrintQuality-Konstanten
Private Const DMRES_DRAFT          As Long = -1 ' Entwurf
Private Const DMRES_LOW            As Long = -2 '  niedrig
Private Const DMRES_MEDIUM         As Long = -3 ' mittel
Private Const DMRES_HIGH           As Long = -4  ' hoch

Public Enum EPrintQuality
    pqDraft = 1
    pqLow = 2
    pqMedium = 3
    pqHigh = 4
End Enum

' DEVMODE dmColor-Konstanten
Private Const DMCOLOR_MONOCHROME   As Long = 1  ' Gerät unterstützt keine Schwarzweißausgabe
Private Const DMCOLOR_COLOR        As Long = 2 ' Gerät unterstützt keine Farbausgabe

' DEVMODE dmDuplex-Konstanten
Private Const DMDUP_SIMPLEX        As Long = 1 ' einseitiges Drucken
Private Const DMDUP_VERTICAL       As Long = 2  ' beidseitig Drucken bei vertikalem Seitenwechsel
Private Const DMDUP_HORIZONTAL     As Long = 3 ' beidseitig Drucken bei horizontalem Seitenwechsel

' DEVMODE dmTTOption-Konstanten
Private Const DMTT_BITMAP          As Long = 1 ' druckt TrueType-Fonts als Grafiken
Private Const DMTT_DOWNLOAD        As Long = 2 ' druckt Downloaded TrueType-Fonts als Soft-Fonts
Private Const DMTT_SUBDEV          As Long = 4 ' druckt Geräte-Fonts statt TrueType-Fonts

' DEVMODE dmCollate-Konstanten
Private Const DMCOLLATE_FALSE      As Long = 0  ' Drucker kann keine Seiten sortieren
Private Const DMCOLLATE_TRUE       As Long = 1  ' Drucker kann Seiten sortieren

' DEVMODE dmDisplayFlags-Konstanten
Private Const DM_GRAYSCALE         As Long = 1  ' Gerät unterstützt keine Farben, Grautöne werden unterstützt
Private Const DM_INTERLACED        As Long = 2 ' Gerät unterstützt Farben

' DEVMODE dmICMMethod-Konstanten
Private Const DMICMMETHOD_NONE     As Long = 1 ' ICM ist abgeschaltet
Private Const DMICMMETHOD_SYSTEM   As Long = 2  ' ICM wird von Windows gesteuert
Private Const DMICMMETHOD_DRIVER   As Long = 3 ' ICM wird vom Treiber gesteuert
Private Const DMICMMETHOD_DEVICE   As Long = 4  ' ICM wird vom Gerät gesteuert

' DEVMODE dmICMIntent-Konstanten
Private Const DMICM_SATURATE       As Long = 1 ' Sättigung wird optimiert
Private Const DMICM_CONTRAST       As Long = 2 ' Kontrast wird optimiert
Private Const DMICM_COLORMETRIC    As Long = 3 ' exakte Farben

' DEVMODE dmMediaType-Konstanten
Private Const DMMEDIA_STANDARD     As Long = 1  ' Standardpapier
Private Const DMMEDIA_GLOSSY       As Long = 2 ' Fotopapier
Private Const DMMEDIA_TRANSPARECNY As Long = 3  ' Folie

' DEVMODE dmDitherType-Konstanten
Private Const DMDITHER_NONE        As Long = 1 ' kein Dithering
Private Const DMDITHER_COARSE      As Long = 2 ' Dithering mit grobem Brush
Private Const DMDITHER_FINE        As Long = 3 ' Dithering mit feinem Brush
Private Const DMDITHER_LINEART     As Long = 4 ' Linien mit schwarz, weiß und grau
Private Const DMDITHER_GRAYSCALE   As Long = 5  ' Grauskala

' GlobalAlloc wFlags-Konstanten
Private Const GHND                 As Long = &H40 ' Kombination von GMEM_MOVEABLE mit GMEM_ZEROINIT
Private Const GMEM_DDESHARE        As Long = &H2000 ' optimiert den Speicher mit der DDE
Private Const GMEM_DISCARDABLE     As Long = &H100  ' Speicher kann überschrieben werden, kann
' nicht mit GMEM_FIXED kombiniert werden
Private Const GMEM_FIXED           As Long = &H0 ' fixiert den Speicher und die Funktion gibt einen
' Pointer zurück
Private Const GMEM_MOVEABLE        As Long = &H2 ' Der Speicher ist beweglich
Private Const GMEM_NOCOMPACT       As Long = &H10 ' Speicher wird nicht minimiert oder andere
' Blöcke für ihn überschrieben
Private Const GMEM_NODISCARD       As Long = &H20 ' überschreibt keine überschreibbaren Blöcke um
' den Speicher zu reservieren
Private Const GMEM_ZEROINIT        As Long = &H40  ' initialisiert den Inhalt des Speicherblocks bei 0
Private Const GPTR                 As Long = &H42 ' kombiniert GMEM_FIXED und GEMEM_ZEROINIT

#If IsUnicode Then
    Private Declare Function PrintDlg Lib "comdlg32" Alias "PrintDlgW" (pPrintDlg As TPrintDialog) As Long
#Else
    Private Declare Function PrintDlg Lib "comdlg32" Alias "PrintDlgA" (pPrintDlg As TPrintDialog) As Long
#End If
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal BytLen As Long)

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long

Private Type TPrintDialog
    lStructSize      As Long        '  4
    hwndOwner        As LongPtr     '  4   8
    hDevMode         As LongPtr     '  4   8
    hDevNames        As LongPtr     '  4   8
    hdc              As LongPtr     '  4   8
    flags            As Long        '  4
    nFromPage        As Integer     '  2
    nToPage          As Integer     '  2
    nMinPage         As Integer     '  2
    nMaxPage         As Integer     '  2
    nCopies          As Integer     '  2
    hInstance        As LongPtr     '  4   8
    lCustData        As Long        '  4
    lpfnPrintHook    As LongPtr     '  4   8
    lpfnSetupHook    As LongPtr     '  4   8
    lpPrintTemplName As LongPtr  ' String '  4   8
    lpSetupTemplName As LongPtr  ' String '  4   8
    hPrintTemplate   As LongPtr     '  4   8
    hSetupTemplate   As LongPtr     '  4   8
End Type                      ' Sum: 108
 
Private Type TDeviceNames
    wDriverOffset    As Integer      '   2
    wDeviceOffset    As Integer      '   2
    wOutputOffset    As Integer      '   2
    wDefault         As Integer      '   2
    'extra(0 To 99)   As Byte 'String * 100
    extra            As String * 100 ' 100
End Type                        ' Sum: 108
 
Private Type TDeviceMode
    dmDeviceName     As String * 32 ' 32
    dmSpecVersion    As Integer     '  2
    dmDriverVersion  As Integer     '  2
    dmSize           As Integer     '  2
    dmDriverExtra    As Integer     '  2
    dmFields         As Long        '  4
    dmOrientation    As Integer     '  2
    dmPaperSize      As Integer     '  2
    dmPaperLength    As Integer     '  2
    dmPaperWidth     As Integer     '  2
    dmScale          As Integer     '  2
    dmCopies         As Integer     '  2
    dmDefaultSource  As Integer     '  2
    dmPrintQuality   As Integer     '  2
    dmColor          As Integer     '  2
    dmDuplex         As Integer     '  2
    dmYResolution    As Integer     '  2
    dmTTOption       As Integer     '  2
    dmCollate        As Integer     '  2
    dmFormName       As String * 32 ' 32
    dmUnusedPadding  As Integer     '  2
    dmBitsPerPixel   As Integer     '  2
    dmPelsWidth      As Long        '  4
    dmPelsHeight     As Long        '  4
    dmDisplayFlags   As Long        '  4
    dmDisplayFrequency As Long      '  4
    dmICMMethod      As Long        '  4
    dmICMIntent      As Long        '  4
    dmMediaType      As Long        '  4
    dmDitherType     As Long        '  4
    dmReserved1      As Long        '  4
    dmReserved2      As Long        '  4
    ' der folgende Abschnitt ist nur für Windows 2000 relevant
    dmPanningWidth   As Long        '  4
    dmPanningHeight  As Long        '  4
End Type                      ' Sum: 154

Private m_PrintDialog As TPrintDialog
Private m_DeviceMode  As TDeviceMode
Private m_DeviceNames As TDeviceNames

' zeigt einen voreingestellten Drucker-Dialog an und wertet diesen aus
Public Function ShowDialog(FOwner As Form) As VbMsgBoxResult
    
    ' Drucker festlegen, der standardmässig ausgewählt sein soll
    With m_DeviceMode
        .dmSize = Len(m_DeviceMode)
        .dmDeviceName = Printer.DeviceName
        .dmFields = DM_PRINTQUALITY
    End With
    
    ' Speicherblock erstellen und Dev_Mode hineinkopieren
    Dim hMemMode As Long: hMemMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(m_DeviceMode)) 'Len(Dev_Mode)=154
    ' Speicherblock erstellen
    Dim pMemMode  As Long: pMemMode = GlobalLock(hMemMode) ' Speicherblock sperren und Pointer ermitteln
    RtlMoveMemory ByVal pMemMode, m_DeviceMode, Len(m_DeviceMode)  ' Struktur in den Speicherblock kopieren
    GlobalUnlock hMemMode ' Speicherblock entsperren
    
    ' Drucker-Eigenschaften zur Überprüfung festlegen
    With m_DeviceNames
        .wDriverOffset = Len(Printer.DriverName)
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        Dim s As String: s = Printer.DriverName & vbNullChar & Printer.DeviceName & vbNullChar & Printer.Port & vbNullChar
        's = StrConv(s, vbNarrow)
        
        .extra = s
        'RtlMoveMemory VarPtr(Dev_Names) + 8, StrPtr(s), Len(s)
    End With
    
    ' Speicherblock erstellen und Dev_Names hineinkopieren
    Dim hMemName As Long: hMemName = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(m_DeviceNames))
    ' Speicherblock erstellen
    Dim pMemName As Long: pMemName = GlobalLock(hMemName) ' Speicherblock sperren und Pointer ermitteln
    RtlMoveMemory ByVal pMemName, m_DeviceNames, Len(m_DeviceNames) ' Struktur in den Speicherblock kopieren
    GlobalUnlock hMemName ' Speicherblock entsperren
    
    ' Dialog-Eigenschaften setzen
    With m_PrintDialog
        .lStructSize = Len(m_PrintDialog) ' 66
        .hInstance = App.hInstance
        .hwndOwner = FOwner.hWnd
        .hDevMode = hMemMode
        .hDevNames = hMemName
        .flags = PD_DISABLEPRINTTOFILE Or PD_RETURNDC Or PD_ALLPAGES Or PD_PAGENUMS Or PD_USEDEVMODECOPIESANDCOLLATE
        '.nFromPage = 1
        '.nToPage = 15
        '.nMinPage = 1
        '.nMaxPage = 15
    End With
    
    ' Dialog aufrufen
    Dim RetVal As Long: RetVal = PrintDlg(m_PrintDialog)
    If RetVal Then
        ShowDialog = vbOK
        ' Dev_Mode wieder aus dem Speicherblock in die Struktur kopieren und
        ' Speicherblock zerstören
        pMemMode = GlobalLock(m_PrintDialog.hDevMode) ' Speicherblock sperren
        RtlMoveMemory m_DeviceMode, ByVal pMemMode, Len(m_DeviceMode) ' Daten aus dem
        ' Speicherblock in die Struktur kopieren
        GlobalUnlock m_PrintDialog.hDevMode ' Speicherblock entsperren
        'GlobalFree PD.hDevMode ' Speicherblock zerstören
        
        
        ' Dev_Names wieder aus dem Speicherblock in die Struktur kopieren und
        ' Speicherblock zerstören
        pMemName = GlobalLock(m_PrintDialog.hDevNames)  ' Speicherblock sperren
        RtlMoveMemory m_DeviceNames, ByVal pMemName, Len(m_DeviceNames) ' Daten aus dem
        ' Speicherblock in die Struktur kopieren
        GlobalUnlock m_PrintDialog.hDevNames ' Speicherblock entsperren
        'GlobalFree PD.hDevNames ' Speicherblock zerstören
            
    Else 'If RetVal = 0 Then
        'MsgBox "Es ist ein Fehler aufgetreten oder Sie haben auf Abbrechen gedrückt.", vbInformation, "Aufrufen des Dialoges"
        ShowDialog = vbCancel
    End If
    GlobalFree m_PrintDialog.hDevMode ' Speicherblock zerstören
    GlobalFree m_PrintDialog.hDevNames ' Speicherblock zerstören
    
    ' Device-Kontext freigeben
    DeleteObject m_PrintDialog.hdc
    
End Function

Public Property Get PrinterName() As String
    Dim s As String
#If IsUnicode Then
    s = Trim$(StrConv(m_DeviceMode.dmDeviceName, vbFromUnicode))
#Else
    s = Trim$(m_DeviceMode.dmDeviceName)
#End If
    s = Left(s, Len(s) - 1)
    PrinterName = s
End Property

Public Property Get FromPage() As Integer
    FromPage = m_PrintDialog.nFromPage
End Property
Public Property Let FromPage(ByVal Value As Integer)
    m_PrintDialog.nFromPage = Value
End Property

Public Property Get ToPage() As Integer
    ToPage = m_PrintDialog.nToPage
End Property
Public Property Let ToPage(ByVal Value As Integer)
    m_PrintDialog.nToPage = Value
End Property

Public Property Get CountCopies() As Integer
    CountCopies = m_DeviceMode.dmCopies
End Property

Public Property Get DoPageSorting() As Boolean
    DoPageSorting = CBool(m_DeviceMode.dmCollate)
End Property

Public Property Get PrintQuality() As EPrintQuality
    PrintQuality = Abs(m_DeviceMode.dmPrintQuality)
End Property

Private Function EPrintQualityToStr(e As EPrintQuality) As String
    Dim s As String
    Select Case e
    Case pqDraft:  s = "Draft"
    Case pqLow:    s = "Low"
    Case pqMedium: s = "Medium"
    Case pqHigh:   s = "High"
    End Select
    EPrintQualityToStr = s
End Function

Public Function PrintQuality_ToStr() As String
    PrintQuality_ToStr = EPrintQualityToStr(Me.PrintQuality)
End Function

Public Function ToStr() As String
    Dim s As String
    s = s & "PrinterName  : " & Me.PrinterName & vbCrLf
    s = s & "FromPage     : " & Me.FromPage & vbCrLf
    s = s & "ToPage       : " & Me.ToPage & vbCrLf
    s = s & "CountCopies  : " & Me.CountCopies & vbCrLf
    s = s & "PrintQuality : " & Me.PrintQuality_ToStr & vbCrLf
    
    ToStr = s
End Function
