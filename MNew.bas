Attribute VB_Name = "MNew"
Option Explicit

Public Function Action(Obj As Object, ByVal ActionName As String) As Action
    Set Action = New Action: Action.New_ Obj, ActionName
End Function
Public Function Action1(Obj As Object, ByVal ActionName As String) As Action1
    Set Action1 = New Action1: Action1.New_ Obj, ActionName
End Function
Public Function Action2(Obj As Object, ByVal ActionName As String) As Action2
    Set Action2 = New Action2: Action2.New_ Obj, ActionName
End Function
Public Function Action3(Obj As Object, ByVal ActionName As String) As Action3
    Set Action3 = New Action3: Action3.New_ Obj, ActionName
End Function
Public Function Action4(Obj As Object, ByVal ActionName As String) As Action4
    Set Action4 = New Action4: Action4.New_ Obj, ActionName
End Function
Public Function Action5(Obj As Object, ByVal ActionName As String) As Action5
    Set Action5 = New Action5: Action5.New_ Obj, ActionName
End Function
Public Function Action6(Obj As Object, ByVal ActionName As String) As Action6
    Set Action6 = New Action6: Action6.New_ Obj, ActionName
End Function
Public Function Action7(Obj As Object, ByVal ActionName As String) As Action7
    Set Action7 = New Action7: Action7.New_ Obj, ActionName
End Function
Public Function Action8(Obj As Object, ByVal ActionName As String) As Action8
    Set Action8 = New Action8: Action8.New_ Obj, ActionName
End Function
Public Function ActionOfString(Obj As Object, ByVal ActionName As String) As ActionOfString
    Set ActionOfString = New ActionOfString: ActionOfString.New_ Obj, ActionName
End Function
Public Function ActionEvent(src As Object, Optional ByVal aID As Long, Optional ByVal aCmd As String, Optional ByVal lwhen As Long, Optional ByVal modifi As Long) As ActionEvent
    Set ActionEvent = New ActionEvent: ActionEvent.New_ src, aID, aCmd, lwhen, modifi
End Function

Public Function AlphaPB(ForePB As PictureBox, BackPB As PictureBox) As AlphaPB
    Set AlphaPB = New AlphaPB: AlphaPB.New_ ForePB, BackPB
End Function

Public Function Aluminium(ByVal BezNum As String, ByVal BezChem As String, ByVal BestKlasse As String, ByVal Erzeugnis As EAluErzeugnisForm, ByVal WerkstZust As String, ByVal Dicke_t_von As Double, ByVal Dicke_t_bis As Double, _
                          ByVal fo As Double, ByVal fu As Double, ByVal A50 As Double, ByVal fohaz As Double, ByVal fuhaz As Double, ByVal rhoohaz As Double, ByVal rhouhaz As Double, ByVal bc As String, ByVal np As Double) As Aluminium
    Set Aluminium = New Aluminium: Aluminium.New_ BezNum, BezChem, BestKlasse, Erzeugnis, WerkstZust, Dicke_t_von, Dicke_t_bis, fo, fu, A50, fohaz, fuhaz, rhoohaz, rhouhaz, bc, np
End Function

Public Function Angle(ByVal angleInRadians As Double) As Angle
    Set Angle = New Angle: Angle.New_ angleInRadians
End Function
Public Function AngleD(ByVal angleInDegrees As Double) As Angle
    Set AngleD = New Angle: AngleD.NewD_ angleInDegrees
End Function
Public Function AngleDec(ByVal angleInRadians) As AngleDec
    Set AngleDec = New AngleDec: AngleDec.New_ angleInRadians
End Function
Public Function AngleDecD(ByVal angleInDegrees) As AngleDec
    Set AngleDecD = New AngleDec: AngleDecD.NewD_ angleInDegrees
End Function
Public Function AngleDecDMS(ByVal deg As Long, ByVal Min As Long, ByVal sec As Double) As AngleDec
    Set AngleDecDMS = New AngleDec: AngleDecDMS.NewDMS_ deg, Min, sec
End Function
Public Function AngleDecG(ByVal angleInGon As Double) As AngleDec
    Set AngleDecG = New AngleDec: AngleDecG.NewG_ angleInGon
End Function
Public Function AngleDecS(ByVal angleInDMS As String) As AngleDec
    Set AngleDecS = New AngleDec: AngleDecS.Parse angleInDMS
End Function
Public Function AngleDeg(ByVal aValueDeg As Double) As Angle
    Set AngleDeg = New Angle: AngleDeg.NewD aValueDeg
End Function
'Public Function AngleDMS(ByVal aDeg As Integer, ByVal aMins As Byte, ByVal aSecs As Byte) As Angle
'    Set AngleDMS = New Angle: AngleDMS.NewDMS aDeg, aMins, aSecs
'End Function
Public Function AngleDMS(ByVal aDeg As Long, ByVal aMin As Double, ByVal asec As Double) As Angle
    Set AngleDMS = New Angle: AngleDMS.NewDMS_ aDeg, aMin, asec
End Function
Public Function AngleG(ByVal angleInGon As Double) As Angle
    Set AngleG = New Angle: AngleG.NewG_ angleInGon
End Function
Public Function AngleS(ByVal angleInDMS As String) As Angle
    Set AngleS = New Angle: AngleS.Parse angleInDMS
End Function

Public Function BBAN(BBANInfo As BBANInfo, sBBAN As String) As BBAN
    Set BBAN = New BBAN: BBAN.New_ BBANInfo, sBBAN
End Function
Public Function BBANInfo(ByVal infoW As String) As BBANInfo
    Set BBANInfo = New BBANInfo: BBANInfo.New_ infoW
End Function
Public Function BBANPart(aChar As String, ByVal aPos As Byte, ByVal aLength As Byte) As BBANPart
    Set BBANPart = New BBANPart: BBANPart.New_ aChar, aPos, aLength
End Function
Public Function BBANValue(bbp As BBANPart, ByVal Value As String) As BBANValue
    Set BBANValue = New BBANValue: BBANValue.New_ bbp, Value
End Function

Public Function BCAlgoProv(ByVal Algo As EBCryptHashAlgo) As BCAlgoProv
    Set BCAlgoProv = New BCAlgoProv: BCAlgoProv.New_ Algo
End Function

Public Function BCHash(ByVal alg As EBCryptHashAlgo) As BCHash
    Set BCHash = New BCHash: BCHash.New_ alg
End Function

'##############################'    Beams    '##############################'
Public Function Beam(aFrm As Form, aParent As PictureBox, aName As String, Index As Long, AColor As Long) As Beam
    Set Beam = New Beam: Beam.New_ aFrm, aParent, aName, Index, AColor
End Function
Public Function Beams(aPB As PictureBox) As Beams
    Set Beams = New Beams: Beams.New_ aPB
End Function

Public Function Beton(ByVal f_ck As Double, ByVal f_ck_cube As Double, Optional ByVal IsLeichtBeton As Boolean = False) As Beton
    Set Beton = New Beton: Beton.New_ f_ck, f_ck_cube, IsLeichtBeton
End Function

Public Function BinaryReader(aFS As FileStream) As BinaryReader
    Set BinaryReader = New BinaryReader: BinaryReader.New_ aFS
End Function
Public Function BinaryWriter(aFS As FileStream) As BinaryWriter
    Set BinaryWriter = New BinaryWriter: BinaryWriter.New_ aFS
End Function

Public Function Bitmap(ByVal aPFN As String) As Bitmap
    Set Bitmap = New Bitmap: Bitmap.New_ aPFN
End Function
Public Function BitmapWH(ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As EPixelFormat) As Bitmap
    Set BitmapWH = New Bitmap: BitmapWH.NewWH Width, Height, PixelFormat
End Function

Public Function BitmapSP(aStdPicture As StdPicture, ByVal aPFN As String) As Bitmap
    Set BitmapSP = New Bitmap: BitmapSP.NewSP aStdPicture, aPFN
End Function

Public Function BlzBic(sLine As String) As BlzBic
    Set BlzBic = New BlzBic: BlzBic.ParseLine sLine
End Function

Public Function BlzBics(aFNm As String) As BlzBics
    Set BlzBics = New BlzBics: BlzBics.ParseFile aFNm
End Function

Public Function Boden(ByVal Name As String, ByVal Homogenbereich As String, ByVal Hauptbestandteil As EBodenHauptbestandteil, ByVal Nebenbestandteil As EBodenNebenbestandteil, ByVal NebenAnteil As EBodenNebenAnteil, ByVal Konsistenz As EBodenKonsistenz, ByVal Stufung As EBodenStufung, _
                      ByVal phi_k_deg As MatBereich, ByVal alp_deg As Double, ByVal bet_deg As Double, ByVal del_deg As Double, ByVal c_k As MatBereich, ByVal cu_k As MatBereich, ByVal Es_k As MatBereich, ByVal gam As MatBereich, ByVal gam_ua As MatBereich, ByVal k_wasser As MatBereich, _
                      ByVal ks_bettung As MatBereich, ByVal qc_spitzendruck As MatBereich, ByVal EModul As Double, ByVal GModul As Double, ByVal KModul As Double, ByVal Nue As Double, ByVal Dichte As Double) As Boden
    Set Boden = New Boden: Boden.New_ Name, Homogenbereich, Hauptbestandteil, Nebenbestandteil, NebenAnteil, Konsistenz, Stufung, phi_k_deg, alp_deg, bet_deg, del_deg, c_k, cu_k, Es_k, gam, gam_ua, k_wasser, ks_bettung, qc_spitzendruck, EModul, GModul, KModul, Nue, Dichte
End Function

Public Function Building(ByVal Name As String, ByVal gps As String) As Building
    Set Building = New Building: Building.New_ Name, gps
End Function

Public Function CalcController(aView As CalcView, aModel As CalcModel) As CalcController
    Set CalcController = New CalcController: CalcController.New_ aView, aModel
End Function

Public Function CalcListener(aCC As CalcController) As CalcListener
    Set CalcListener = New CalcListener: CalcListener.New_ aCC
End Function

Public Function Calendar(ByVal ThisYear As CalendarYear, Optional ByVal LastYear As CalendarYear = Nothing, _
                                                         Optional ByVal NextYear As CalendarYear = Nothing) As Calendar
    Set Calendar = New Calendar: Calendar.New_ ThisYear, LastYear, NextYear
End Function
Public Function CalendarDay(ByVal Month As CalendarMonth, ByVal DayInMonth As Integer) As CalendarDay
    Set CalendarDay = New CalendarDay: CalendarDay.New_ Month, DayInMonth
End Function
Public Function CalendarDefault() As Calendar
    Dim thisY As Integer: thisY = GetDefaultYear
    Set CalendarDefault = MNew.Calendar(MNew.CalendarYear(thisY, 1, 12), _
                                        MNew.CalendarYear(thisY - 1, 12, 12), _
                                        MNew.CalendarYear(thisY + 1, 1, 1))
End Function
Public Function CalendarEvent(ByVal EventFromDate As Date, ByVal EventToDate As Date, ByVal EventName As String) As CalendarEvent
    Set CalendarEvent = New CalendarEvent: CalendarEvent.New_ EventFromDate, EventToDate, EventName
End Function
Public Function CalendarMonth(CalYear As CalendarYear, ByVal Month As Integer) As CalendarMonth
    Set CalendarMonth = New CalendarMonth: CalendarMonth.New_ CalYear, Month
End Function
Public Function CalendarView(aCalendar As Calendar, Canvas As PictureBox) As CalendarView
    Set CalendarView = New CalendarView: CalendarView.New_ aCalendar, Canvas
End Function
Public Function CalendarYear(ByVal Year As Integer, Optional ByVal StartMonth As Integer = 1, _
                                                    Optional ByVal EndMonth As Integer = 12) As CalendarYear
    Set CalendarYear = New CalendarYear: CalendarYear.New_ Year, StartMonth, EndMonth
End Function

Public Function CBElement(aCBFormat As ClipboardFormat, BytArr() As Byte) As CBElement
    Set CBElement = New CBElement: CBElement.New_ aCBFormat, BytArr
End Function

Public Function CCollection(ByVal IsHashed As Boolean, Optional Col As Collection = Nothing, Optional Name As String = "", Optional ByVal OptionBaseLBound As Long = 1) As CCollection
    Set CCollection = New CCollection: CCollection.New_ IsHashed, Col, Name, OptionBaseLBound
End Function

Public Sub CDeserialize(aBR As BinaryReader, aObj As ISerializable, Optional ByVal aOffSet As Long)
    Call aObj.Deserialize(aBR, aOffSet)
End Sub

Public Function CholeskyDecomposition(Value As Matrix) As CholeskyDecomposition
    Set CholeskyDecomposition = New CholeskyDecomposition: CholeskyDecomposition.New_ Value
End Function

Public Function City(ByVal Name As String, ByVal PostalCode As String) As City
    Set City = New City: City.New_ Name, PostalCode
End Function

Public Function CMatOp(ByVal P As Long, s As String, Optional ByVal nk As Byte) As CMatOp
    Set CMatOp = New CMatOp: CMatOp.New_ P, s, nk
End Function
Public Function CMatOpByVal(ByVal pDbl As Long, Optional ByVal mRows As Byte = 1, Optional ByVal nCols As Byte = 1) As CMatOp
    Set CMatOpByVal = New CMatOp: CMatOpByVal.NewByVal pDbl, mRows, nCols
End Function
Public Function CMatOpByValQ(ByVal pDbl As Long, Optional ByVal maxRowsCols As Byte = 1) As CMatOp
    Set CMatOpByValQ = New CMatOp: CMatOpByValQ.NewByVal pDbl, maxRowsCols, maxRowsCols
End Function

Public Function CoderVB(Optional ByVal WriteClsHeader As Boolean = False) As CoderVB
    Set CoderVB = New CoderVB: CoderVB.New_ WriteClsHeader
End Function

Public Function ColorSelector(aTimer As Timer, aButton As CommandButton, aColorView As PictureBox, aLabel As Label) As ColorSelector
    Set ColorSelector = New ColorSelector: ColorSelector.New_ aTimer, aButton, aColorView, aLabel
End Function

Public Function ComInterface(aIID As GUID, Optional aName As String) As ComInterface
    Set ComInterface = New ComInterface: ComInterface.New_ aIID, aName
End Function
Public Function ComIntfcFunction(aComIntfc As ComInterface, ByVal aName As String, ByVal aIndex As Long, ByVal ReturnType As EVbVarType) As ComIntfcFunction
    Set ComIntfcFunction = New ComIntfcFunction: ComIntfcFunction.New_ aComIntfc, aName, aIndex, ReturnType
End Function

Public Function ConfigIniDocument(pfn As PathFileName) As ConfigIniDocument
    Set ConfigIniDocument = New ConfigIniDocument: ConfigIniDocument.New_ pfn
End Function
Public Function ConfigIniKeyValue(IniFile As ConfigIniDocument, _
                                  Section As ConfigIniSection, _
                                  keyName As String, _
                                  Optional VarDefault As Variant) As ConfigIniKeyValue
    Set ConfigIniKeyValue = New ConfigIniKeyValue: ConfigIniKeyValue.New_ IniFile, Section, keyName, VarDefault
End Function
Public Function ConfigIniKeyValueS(IniFile As ConfigIniDocument, _
                                   SectionName As String, _
                                   keyName As String, _
                                   Optional VarDefault As Variant) As ConfigIniKeyValue
    Set ConfigIniKeyValueS = New ConfigIniKeyValue: ConfigIniKeyValueS.NewS IniFile, SectionName, keyName, VarDefault
End Function
Public Function ConfigIniSection(IniFile As ConfigIniDocument, SectionName As String) As ConfigIniSection
    Set ConfigIniSection = New ConfigIniSection: ConfigIniSection.New_ IniFile, SectionName
End Function

Public Function CPlayerButtons(ByVal FastRewind As OptionButton, ByVal PlayPause As CommandButton, ByVal AllStop As CommandButton, ByVal FastForward As OptionButton, ByVal Timer As Timer, ByVal Label As Label, _
                               Optional SongFirst As CommandButton, Optional SongPrev As CommandButton, Optional SongNext As CommandButton, Optional SongLast As CommandButton, Optional ByVal SongNr As Long = 0) As CPlayerButtons
    Set CPlayerButtons = New CPlayerButtons: CPlayerButtons.New_ FastRewind, PlayPause, AllStop, FastForward, Timer, Label, SongFirst, SongPrev, SongNext, SongLast, SongNr, MinSongNr, MaxSongNr
End Function

Public Function CPUClock(ByVal Rate As Double) As CPUClock
    Set CPUClock = New CPUClock: CPUClock.New_ Rate
End Function

Public Function CRC32(aHashAlgo As ECRC32Algo) As CRC32
    Set CRC32 = New CRC32: CRC32.New_ aHashAlgo
End Function

Public Function CreateNewCustomer() As DemoCustomer
    Set CreateNewCustomer = New DemoCustomer
End Function

Public Sub CSerialize(aBW As BinaryWriter, aObj As ISerializable, Optional ByVal aOffSet As Long)
    Call aObj.Serialize(aBW, aOffSet)
End Sub

Public Function CSHA256(aHashAlgo As ECRC32Algo) As CSHA256
    Set CSHA256 = New CSHA256 ': CSHA.New_ aHashAlgo
End Function

Public Function CTwig(aLeaf As Leaf) As Twig
    'a simple cast
    Set CTwig = aLeaf
End Function

Public Function CultureInfo(ByVal aLCID As Long) As CultureInfo
    Set CultureInfo = New CultureInfo: CultureInfo.New_ aLCID
End Function

Public Function CultureInfoN(ByVal aLcidName As String) As CultureInfo
    Set CultureInfoN = New CultureInfo: CultureInfoN.NewN aLcidName
End Function

Public Function DBcrud(Col As CCollection, ByVal IsHashed As Boolean, ListBox, _
                       BtnAdd As CommandButton, Optional BtnAddClone, Optional BtnInsert, Optional BtnInsertClone, Optional BtnEdit, _
                       Optional BtnDelete, Optional BtnMoveUp, Optional BtnMoveDown, Optional BtnSortUp, Optional BtnSortDown, Optional BtnSearch, Optional TxtSearch) As DBcrud
    Set DBcrud = New DBcrud: DBcrud.New_ Col, ListBox, BtnAdd, BtnAddClone, BtnInsert, BtnInsertClone, BtnEdit, BtnDelete, BtnMoveUp, BtnMoveDown, BtnSortUp, BtnSortDown, BtnSearch, TxtSearch
End Function

Public Function DisplayConnector(ByVal DevName As String, ByRef i_inout As Long) As DisplayConnector
    Set DisplayConnector = New DisplayConnector: DisplayConnector.New_ DevName, i_inout
End Function

Public Function DisplayDevice(Optional ByVal DevName As String = vbNullString, Optional ByRef i_inout As Long = 0) As DisplayDevice
    Set DisplayDevice = New DisplayDevice: DisplayDevice.New_ DevName, i_inout
End Function
Public Function DisplayModeFilter(ListOfDisplayModes As Collection) As DisplayModeFilter
    Set DisplayModeFilter = New DisplayModeFilter: DisplayModeFilter.New_ ListOfDisplayModes
End Function

Public Function Document(aIPBase As IPAddress, aLastIP As IPAddress, nSearchIPs As Long, aScanner As IPPingScanner) As Document
    Set Document = New Document: Document.New_ aIPBase, aLastIP, nSearchIPs, aScanner
End Function
Public Function DocumentP(aPFN As String) As Document
    Set DocumentP = New Document: DocumentP.New_ aPFN
End Function
' ############################## '    DocumentPStore    ' ############################## '
Public Function DocumentPStore(aPFN As PathFileName, PossFileProps As List) As DocumentPStore 'PossFileProps As List(Of PropKey))
    Set DocumentPStore = New DocumentPStore: DocumentPStore.New_ aPFN, PossFileProps
End Function
' ############################## '    DocumentTsv    ' ############################## '
'Public Function DocumentTsv(aPFN As String, Optional ByVal aPropertyLists As List = Nothing) As DocumentTsv
'    Set DocumentTsv = New DocumentTsv: DocumentTsv.New_ aPFN, aPropertyLists
'End Function
Public Function DocumentTsv(aPFN As PathFileName, Optional ByVal aPropertyLists As List = Nothing) As DocumentTsv
    Set DocumentTsv = New DocumentTsv: DocumentTsv.New_ aPFN, aPropertyLists
End Function

Public Function DocumentVBAXL(aXLWkbk As PathFileName) As DocumentVBAXL
    Set DocumentVBAXL = New DocumentVBAXL: DocumentVBAXL.New_ aXLWkbk
End Function

Public Function DOMDocument60(Optional bAsync As Boolean = True, Optional doValidateOnParse As Boolean = True, Optional doResolveExternals As Boolean = False, Optional doPreserveWhitespace As Boolean = False) As DOMDocument60
    Set DOMDocument60 = New MSXML2.DOMDocument60
    With DOMDocument60
        .Async = bAsync                            'by default true
        .validateOnParse = doValidateOnParse       'by default true
        .resolveExternals = doResolveExternals     'by default false anyway
        .preserveWhiteSpace = doPreserveWhitespace 'by default false
    End With
End Function

Public Function DOMDocument60F(ByVal aPathFileName As String) As DOMDocument60
    If Len(aPathFileName) = 0 Then Exit Function
    Set DOMDocument60F = MNew.DOMDocument60(bAsync:=False, doValidateOnParse:=False, doResolveExternals:=False, doPreserveWhitespace:=True)
    If Not DOMDocument60F.Load(aPathFileName) Then
        MsgBox "Can't create DOM from " & aPathFileName
        Set DOMDocument60F = Nothing
        Exit Function
    End If
End Function

'Public Function EndianBinaryReader(st As Stream) As EndianBinaryReader
'    Set EndianBinaryReader = New EndianBinaryReader: EndianBinaryReader.New_ st
'End Function
Public Function EBinaryReader(st As Stream) As EBinaryReader
    Set EBinaryReader = New EBinaryReader: EBinaryReader.New_ st
End Function

Public Function EigenvalueDecomposition(Value As Matrix) As EigenvalueDecomposition
    Set EigenvalueDecomposition = New EigenvalueDecomposition:  EigenvalueDecomposition.New_ Value
End Function

Public Function Email(ByVal SenderAddress As String, ReceiverAddress As String) As Email
    Set Email = New Email: Email.New_ SenderAddress, ReceiverAddress
End Function

Public Function EngValue(ByVal Value, ByVal u As EEinheit) As EngValue
    Set EngValue = New EngValue: EngValue.New_ Value, u
End Function

Public Function FestivalDay(ByVal aDate As Date, ByVal EFestival As Long) As FestivalDay
    Set FestivalDay = New FestivalDay: FestivalDay.New_ aDate, EFestival
End Function

Public Function FilEBinReader(PathFileName As String) As FilEBinReader
    Set FilEBinReader = New FilEBinReader: FilEBinReader.New_ PathFileName
End Function

Public Function FileProperty(aPropKey As PropKey, aValue As VVariant) As FileProperty
    Set FileProperty = New FileProperty: FileProperty.New_ aPropKey, aValue
End Function

Public Function FilePropertyPkhe(aPropKeyHentry As PropKeyHEntry) As FileProperty
    Dim vt As EVbVarType:   vt = aPropKeyHentry.EVbVarType
    Set FilePropertyPkhe = MNew.FileProperty(MNew.PropKeyPkhe(aPropKeyHentry), MNew.VVariantVt(vt))
End Function

'##############################'  System.IO   '##############################'
Public Function FileStream(aPathFileName As String, _
                           Optional FMode As FileMode = FileMode_Random, _
                           Optional FAccess As FileAccess = FileAccess_ReadWrite, _
                           Optional FShare As FileShare = FileShare_None) As FileStream
  Set FileStream = New FileStream: FileStream.New_ aPathFileName, FMode, FAccess, FShare
End Function

Public Function FileVersionInfo(aPathFileName As String) As FileVersionInfo
    Set FileVersionInfo = New FileVersionInfo: FileVersionInfo.New_ aPathFileName
End Function
'Public Function FileVersionInfo(aFileName As String) As FileVersionInfo
'    Set FileVersionInfo = New FileVersionInfo ': FileVersionInfo.NewC aFileName
'    Set FileVersionInfo = FileVersionInfo.GetVersionInfo(aFileName)
'End Function

Public Function FocusRect(ByVal aHDC As LongPtr) As FocusRect
    Set FocusRect = New FocusRect: FocusRect.New_ aHDC
End Function

Public Function FormatRange(ByVal vt As VbVarType, ByVal Min As Double, ByVal Max As Double) As String
    Dim s As String
    Select Case vt
    Case VbVarType.vbDate
        Dim datMin As Date: datMin = Min
        Dim datMax As Date: datMax = Max
        s = datMin & " | " & datMax
    Case Else
        s = Min & " | " & Max
    End Select
    FormatRange = s
End Function
'ist Shared muﬂ deshalb in ein Modul
'von Bitmap:
'Public Shared Function FromResource(ByVal hinstance As System.IntPtr, ByVal bitmapName As String) As System.Drawing.Bitmap
Public Function FromResource(Optional ByVal hhinstance As Long, Optional ByVal bitmapNameResID As String) As Image 'System.Drawing.Bitmap
    Set FromResource = New Image: FromResource.FromStdPicture LoadResPicture(CLng(bitmapNameResID), vbResBitmap)
End Function

Public Function Func(Obj As Object, ByVal FuncName As String) As Func
    Set Func = New Func: Func.New_ Obj, FuncName
End Function
Public Function Func1(Obj As Object, ByVal FuncName As String) As Func1
    Set Func1 = New Func1: Func1.New_ Obj, FuncName
End Function
Public Function Func2(Obj As Object, ByVal FuncName As String) As Func2
    Set Func2 = New Func2: Func2.New_ Obj, FuncName
End Function
Public Function Func3(Obj As Object, ByVal FuncName As String) As Func3
    Set Func3 = New Func3: Func3.New_ Obj, FuncName
End Function
Public Function Func4(Obj As Object, ByVal FuncName As String) As Func4
    Set Func4 = New Func4: Func4.New_ Obj, FuncName
End Function
Public Function Func5(Obj As Object, ByVal FuncName As String) As Func5
    Set Func5 = New Func5: Func5.New_ Obj, FuncName
End Function
Public Function Func6(Obj As Object, ByVal FuncName As String) As Func6
    Set Func6 = New Func6: Func6.New_ Obj, FuncName
End Function
Public Function Func7(Obj As Object, ByVal FuncName As String) As Func7
    Set Func7 = New Func7: Func7.New_ Obj, FuncName
End Function
Public Function Func8(Obj As Object, ByVal FuncName As String) As Func8
    Set Func8 = New Func8: Func8.New_ Obj, FuncName
End Function
Public Function FuncOfString(Obj As Object, ByVal FuncName As String) As FuncOfString
    Set FuncOfString = New FuncOfString: FuncOfString.New_ Obj, FuncName
End Function

Public Function FunctionHook(ByVal pFnc As Long) As FunctionHook
    Set FunctionHook = New FunctionHook: FunctionHook.New_ pFnc
End Function

Public Function GDIBackBuffer(aPB) As GDIBackBuffer 'As PictureBox) As GDIBackBuffer
    Set GDIBackBuffer = New GDIBackBuffer: GDIBackBuffer.New_ aPB
End Function

Public Function GDIBitmap(aPFN As String) As GDIBitmap
    Set GDIBitmap = New GDIBitmap: GDIBitmap.New_ aPFN
End Function

Public Function GeoPos(Latitude As AngleDec, Longitude As AngleDec, Optional ByVal Height As Double, Optional ByVal Name As String) As GeoPos
    Set GeoPos = New GeoPos: GeoPos.New_ Latitude, Longitude, Height, Name
End Function
Public Function GeoPosS(s As String) As GeoPos
    Set GeoPosS = New GeoPos: GeoPosS.Parse s
End Function

Public Function GetDefaultYear() As Integer
    Dim dNow  As Date:     dNow = DateTime.Now
    Dim Year  As Integer:  Year = DateTime.Year(dNow)
    Dim Month As Integer: Month = DateTime.Month(dNow)
    Dim Day   As Integer:   Day = DateTime.Day(dNow)
    GetDefaultYear = Year + IIf((Month = 12) And (Day = 15), 1, 0)
End Function

Public Function GitHubUser(ByVal LoginName As String, ByVal Id As Long, ByVal NodeID As String, ByVal AvatarURL As String, ByVal GrAvatarID As String, ByVal Url As String, ByVal HtmlUrl As String, _
                           ByVal FollowersUrl As String, ByVal FollowingUrl As String, ByVal GistsUrl As String, ByVal StarredUrl As String, ByVal SubscriptionsUrl As String, ByVal OrganizationsUrl As String, _
                           ByVal ReposUrl As String, ByVal EventsUrl As String, ByVal ReceivedEventsUrl As String, ByVal UserType As String, ByVal UserViewType As String, ByVal SiteAdmin As Boolean) As GitHubUser
    Set GitHubUser = New GitHubUser: GitHubUser.New_ LoginName, Id, NodeID, AvatarURL, GrAvatarID, Url, HtmlUrl, _
                                                     FollowersUrl, FollowingUrl, GistsUrl, StarredUrl, SubscriptionsUrl, OrganizationsUrl, _
                                                     ReposUrl, EventsUrl, ReceivedEventsUrl, UserType, UserViewType, SiteAdmin
End Function

Public Function GitHubUserD(Dictionary As Dictionary) As GitHubUser
    Dim gu As GitHubUser: Set gu = New GitHubUser: If gu.Read(Dictionary) Then Set GitHubUserD = gu
End Function

Public Function GraphicView(aCanvas As PictureBox, aPoint3DList As List) As GraphicView
    Set GraphicView = New GraphicView: GraphicView.New_ aCanvas, aPoint3DList
End Function

' ############################## '    Guid    ' ############################## '
Public Function GUID(ByVal Data1 As Long, ByVal Data2 As Integer, ByVal Data3 As Integer, _
                     ByVal Data50 As Byte, ByVal Data51 As Byte, ByVal Data52 As Byte, ByVal Data53 As Byte, _
                     ByVal Data54 As Byte, ByVal Data55 As Byte, ByVal Data56 As Byte, ByVal Data57 As Byte) As GUID
    Set GUID = New GUID: GUID.New_ Data1, Data2, Data3, Data50, Data51, Data52, Data53, Data54, Data55, Data56, Data57
End Function
'Public Function GUID(ByVal A As Long, ByVal b As Integer, ByVal c As Integer, _
'                      ByVal d0 As Byte, ByVal d1 As Byte, ByVal d2 As Byte, ByVal d3 As Byte, _
'                      ByVal d4 As Byte, ByVal d5 As Byte, ByVal d6 As Byte, ByVal d7 As Byte) As GUID
'    Set GUID = New GUID: GUID.New_ A, b, c, d0, d1, d2, d3, d4, d5, d6, d7
'End Function
Public Function GuidCo() As GUID
    Set GuidCo = New GUID: GuidCo.NewCo
End Function
Public Function GuidD(Data1 As Long, Data2 As Integer, Data3 As Integer, d() As Byte) As GUID
    Set GuidD = New GUID: GuidD.NewD Data1, Data2, Data3, d
End Function
Public Function Guidlp(ByVal lpGuid As LongPtr) As GUID
    Set Guidlp = New GUID: RtlMoveMemory ByVal Guidlp.ptr, ByVal lpGuid, Guidlp.Size
End Function
Public Function GuidPK(ByVal sGuid As String, ByVal PID As Long) As GUID
    Set GuidPK = New GUID: GuidPK.NewPK sGuid, PID
End Function
Public Function GuidS(s As String) As GUID
    Set GuidS = New GUID
    If Not GuidS.Parse(s) Then Set GuidS = Nothing
End Function

Public Function HeapEntry(aSnapShot As Snapshot, ByVal Index As Long) As HeapEntry
    Set HeapEntry = New HeapEntry: HeapEntry.New_ aSnapShot, Index
End Function
Public Function HeapList(aSnapShot As Snapshot) As HeapList
    Set HeapList = New HeapList: HeapList.New_ aSnapShot
End Function

'HeapSorter
Public Function HeapSorter(aCmp As List) As HeapSorter ', aSwp As Object) As HeapSorter ' As IComparable, ISwappable
    Set HeapSorter = New HeapSorter: HeapSorter.New_ aCmp ', aSwp
End Function

Public Function Holz(ByVal Klasse As EHolzKlasse, ByVal f_mk_Nmm2 As Double, ByVal f_t0k_Nmm2 As Double, ByVal EModul0mean_kNmm2 As Double, ByVal Rohdichterhok_kgm3 As Double) As Holz
    Set Holz = New Holz: Holz.New_ Klasse, f_mk_Nmm2, f_t0k_Nmm2, EModul0mean_kNmm2, Rohdichterhok_kgm3
End Function

Public Function IBAN(IBANInfos As IBANInfos, siban As String) As IBAN
    Set IBAN = New IBAN: IBAN.New_ IBANInfos, siban
End Function

Public Function IBANCreator(IBANInfos As IBANInfos, aIBANInfo As IBANInfo, List As Collection) As IBANCreator 'sArr() As String) As IBANCreator
    Set IBANCreator = New IBANCreator: IBANCreator.New_ IBANInfos, aIBANInfo, List 'sArr
End Function

Public Function IBANInfo(Name As String, LC As String, bbinf As BBANInfo) As IBANInfo
    Set IBANInfo = New IBANInfo: IBANInfo.New_ Name, LC, bbinf
End Function

Public Function IBANInfos(aPFN As String) As IBANInfos
    Set IBANInfos = New IBANInfos: IBANInfos.New_ aPFN
End Function

Public Function IDLInterface(aName As String, aUUID As String, Inherits As String) As IDLInterface
    Set IDLInterface = New IDLInterface: IDLInterface.New_ aName, aUUID, Inherits
End Function

Public Function Image(aPic As StdPicture) As Image
    Set Image = New Image: Image.FromStdPicture aPic
End Function

Public Function IndentStack(aTabSize As Long) As IndentStack
    Set IndentStack = New IndentStack: IndentStack.New_ aTabSize
End Function

Public Function InternetURL(ByVal sURL As String, Optional bufferSize As Long = 2048) As InternetURL
    Set InternetURL = New InternetURL: InternetURL.New_ sURL, bufferSize
End Function

Public Function IPAddress(ByVal i1 As Integer, ByVal i2 As Integer, ByVal i3 As Integer, ByVal i4 As Integer, Optional i5, Optional i6, Optional i7, Optional i8, Optional Port) As IPAddress
    If IsMissing(i5) Then
        Set IPAddress = IPAddressV4(CByte(i1), CByte(i2), CByte(i3), CByte(i4), Port)
    Else
        Set IPAddress = IPAddressV6(i1, i2, i3, i4, i5, i6, i7, i8, Port)
    End If
End Function
Public Function IPAddressV(StrLngBytesNewAddress) As IPAddress
    Set IPAddressV = New IPAddress: IPAddressV.New_ StrLngBytesNewAddress
End Function
Public Function IPAddressV4(ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte, ByVal b4 As Byte, Optional Port) As IPAddress
    Set IPAddressV4 = New IPAddress: IPAddressV4.NewV4 b1, b2, b3, b4, Port
End Function
Public Function IPAddressV6(ByVal i1 As Integer, ByVal i2 As Integer, ByVal i3 As Integer, ByVal i4 As Integer, ByVal i5 As Integer, ByVal i6 As Integer, ByVal i7 As Integer, ByVal i8 As Integer, Optional Port) As IPAddress
    Set IPAddressV6 = New IPAddress: IPAddressV6.NewV6 i1, i2, i3, i4, i5, i6, i7, i8, Port
End Function
Public Function IPAddressV4Rnd() As IPAddress
    Randomize
    Set IPAddressV4Rnd = New IPAddress: IPAddressV4Rnd.NewV4 RndUInt8, RndUInt8, RndUInt8, RndUInt8
End Function
Public Function IPAddressV6Rnd() As IPAddress
    Randomize
    Set IPAddressV6Rnd = New IPAddress: IPAddressV6Rnd.NewV6 RndUInt16, RndUInt16, RndUInt16, RndUInt16, RndUInt16, RndUInt16, RndUInt16, RndUInt16
End Function

Public Function IRange(ByVal MinValue As Double, ByVal MaxValue As Double, ByVal RangeType As ERangeType) As IRange
    Select Case RangeType
    Case ERangeType.ExMinExMax: Set IRange = MNew.RangeExMinExMax(MinValue, MaxValue)
    Case ERangeType.ExMinInMax: Set IRange = MNew.RangeExMinInMax(MinValue, MaxValue)
    Case ERangeType.InMinExMax: Set IRange = MNew.RangeInMinExMax(MinValue, MaxValue)
    Case ERangeType.InMinInMax: Set IRange = MNew.RangeInMinInMax(MinValue, MaxValue)
    End Select
End Function

'Function IsInIDE() As Boolean
'Try: On Error GoTo Catch
'    Debug.Print 1 / 0
'    IsInIDE = False: Exit Function
'Catch: IsInIDE = True
'End Function
Function IsInIDE() As Boolean
    IsInIDE = App.LogMode = 0 'LogModeConstants.vbLogAuto
End Function

Public Function LANAdapter(LanNumber As Long, MACAddress As MACAddress) As LANAdapter
    Set LANAdapter = New LANAdapter: LANAdapter.New_ LanNumber, MACAddress
End Function

'Public Function ListOfLeaf(bUseHashes As Boolean) As ListOfLeaf
'    Set ListOfLeaf = New ListOfLeaf
'    ListOfLeaf.New_ bUseHashes
'End Function
Public Function Leaf(ByVal aName As String, ByVal aValue) As Leaf
    Set Leaf = New Leaf: Leaf.New_ aName, aValue
End Function

' ############################## '    List    ' ############################## '
Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

'not inside of DocumentTsv because FileProperty is unknown in vb-project PReadPropKeyH
Public Function ListOfFileProperty(doc As DocumentTsv) As List '(Of FileProperty)
    Dim lifp As List: Set lifp = MNew.List(EDataType.vbObject)
    Dim i As Long, j As Long, le As List
    Dim pkhe As PropKeyHEntry, fp As FileProperty
    Dim pl As List: Set pl = doc.PropertyLists
    For i = 0 To pl.Count - 1
        Set le = pl.item(i)
        For j = 0 To le.Count - 1
            Set pkhe = le.item(j)
            Set fp = MNew.FilePropertyPkhe(pkhe)
            lifp.Add fp 'pk.ToPossFileProperty
        Next
    Next
    Set ListOfFileProperty = lifp
End Function

Public Function LListBox(aListBox As ListBox) As LListBox
    Set LListBox = New LListBox: LListBox.New_ aListBox
End Function

Public Function LoadFromResource(nID As Long) As IPictureDisp
    LoadFromResource = LoadResPicture(nID, vbResBitmap)
End Function

Function LogModeConstants_ToStr(e As LogModeConstants) As String
    Dim s As String
    Select Case e
    Case LogModeConstants.vbLogAuto:      s = "vbLogAuto"      ' 0
    Case LogModeConstants.vbLogOff:       s = "vbLogOff"       ' 1
    Case LogModeConstants.vbLogToFile:    s = "vbLogToFile"    ' 2
    Case LogModeConstants.vbLogToNT:      s = "vbLogToNT"      ' 3
    Case LogModeConstants.vbLogOverwrite: s = "vbLogOverwrite" '16 &H10
    Case LogModeConstants.vbLogThreadID:  s = "vbLogThreadID"  '32 &H20
    Case Else: s = CStr(e)
    End Select
    LogModeConstants_ToStr = s
End Function

Public Function LuDecomposition(Value As Matrix) As LuDecomposition
    Set LuDecomposition = New LuDecomposition:  LuDecomposition.New_ Value
End Function

Public Function MACAddress(ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte, ByVal b4 As Byte, ByVal b5 As Byte, ByVal b6 As Byte, Optional sep As String = "-") As MACAddress
    Set MACAddress = New MACAddress: MACAddress.New_ b1, b2, b3, b4, b5, b6, sep
End Function

Public Function MACAddressA(bytes05() As Byte, Optional sep As String = "-") As MACAddress
    Set MACAddressA = New MACAddress: MACAddressA.New_ bytes05(0), bytes05(1), bytes05(2), bytes05(3), bytes05(4), bytes05(5), sep
End Function

Public Function MACAddressS(s As String) As MACAddress
    Set MACAddressS = New MACAddress: MACAddressS.NewS s
End Function

Public Function MatBereich(Optional Value, Optional Bereich_von, Optional Bereich_bis, Optional ByVal Auswahl As EBereichAuswahl = EBereichAuswahl.Mittelwert) As MatBereich
    Set MatBereich = New MatBereich: MatBereich.New_ Value, Bereich_von, Bereich_bis, Auswahl
End Function

Public Function MatBereichB(Optional Bereich_von, Optional Bereich_bis, Optional ByVal Auswahl As EBereichAuswahl = EBereichAuswahl.Mittelwert) As MatBereich
    Set MatBereichB = New MatBereich: MatBereich.NewB Bereich_von, Bereich_bis, Auswahl
End Function

Public Function MaterialMC(ByVal EModul As Double, ByVal GModul As Double, ByVal KModul As Double, ByVal Nue As Double, ByVal Dichte_rho As Double, Optional ByVal Wichte_gam As Double = 0, Optional ByVal Name As String = "", Optional ByVal norm As String = "") As MaterialMC
    Set MaterialMC = New MaterialMC: MaterialMC.New_ EModul, GModul, KModul, Nue, Dichte_rho, Wichte_gam, Name, norm
End Function

Function MathFunction(ByVal addrof As LongPtr) As MathFunction
    Set MathFunction = New MathFunction
    Dim pVTable As LongPtr
    GetMem4 ByVal ObjPtr(MathFunction), pVTable
    GetMem4 addrof, ByVal (pVTable + 28)
End Function

Public Function MathFunctionD(aDF As TDelegateFunction, ByVal pFncAddr As LongPtr) As MathFunction 'ICallDoubleReturnDouble
    Set MathFunctionD = New_DelegateFunction(aDF, pFncAddr)
End Function

Public Function MathFunctionView(Canvas As PictureBox) As MathFunctionView
    Set MathFunctionView = New MathFunctionView: MathFunctionView.New_ Canvas
End Function

Public Function Matrix(ByVal Rows As Long, ByVal Columns As Long, Optional ByVal Value) As Matrix
    Set Matrix = New Matrix: Matrix.New_ Rows, Columns, Value
End Function

Public Function MemGlobal(ByVal flags As EMemGlobalFlags, ByVal SizeBytes As Long, Optional ByVal pCopyFrom As LongPtr = 0) As MemGlobal
    Set MemGlobal = New MemGlobal: MemGlobal.New_ flags, SizeBytes, pCopyFrom
End Function

Public Function MemHeap(ByVal flags As EMemGlobalFlags, ByVal SizeBytes As Long, Optional ByVal pCopyFrom As LongPtr = 0) As MemHeap
    Set MemHeap = New MemHeap ': MemHeap.New_ Flags, SizeBytes, pCopyFrom
End Function

Public Function MemLocal(ByVal flags As EMemGlobalFlags, ByVal SizeBytes As Long, Optional ByVal pCopyFrom As LongPtr = 0) As MemLocal
    Set MemLocal = New MemLocal ': MemLocal.New_ Flags, SizeBytes, pCopyFrom
End Function

Public Function MessageDialog(ByVal OwnerHwnd As Long, ByVal Content As String, Optional ByVal Title As String = vbNullString) As MessageDialog
    Set MessageDialog = New MessageDialog: MessageDialog.New_ OwnerHwnd, Content, Title
End Function

'Public Function ModalDialog(aDialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton, Optional ByVal IsReadOnly As Boolean = False) As ModalDialog
'    Set ModalDialog = New ModalDialog: ModalDialog.New_ aDialog, BtnOK, BtnCancel, IsReadOnly
'End Function
'
'Public Function ModalDialog(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton) As ModalDialog
'    Set ModalDialog = New ModalDialog: ModalDialog.New_ Dialog, BtnOK, BtnCancel
'End Function
'Public Function ModalDialog(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton) As ModalDialog
'    Set ModalDialog = New ModalDialog: ModalDialog.New_ Dialog, BtnOK, BtnCancel
'End Function
'Public Function ModalDialog(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton, Optional ByVal IsReadOnly As Boolean = False) As ModalDialog
'    Set ModalDialog = New ModalDialog: ModalDialog.New_ Dialog, BtnOK, BtnCancel, IsReadOnly
'End Function
'Public Function ModalDialog(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton) As ModalDialog
'    Set ModalDialog = New ModalDialog: ModalDialog.New_ Dialog, BtnOK, BtnCancel
'End Function

Public Function ModalDialog(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton, Optional ByVal IsReadOnly As Boolean = False) As ModalDialog
    Set ModalDialog = New ModalDialog: ModalDialog.New_ Dialog, BtnOK, BtnCancel, IsReadOnly
End Function

Public Function ModuleEntry(aSnapShot As Snapshot) As ModuleEntry
    Set ModuleEntry = New ModuleEntry: ModuleEntry.New_ aSnapShot
End Function

Public Function Monitor(ByVal HMONITOR As LongPtr, ByVal HDC As LongPtr, ByVal lpClipRECT As LongPtr) As Monitor
    Set Monitor = New Monitor: Monitor.New_ HMONITOR, HDC, lpClipRECT
End Function
Public Function MonitorP(ByVal X As Long, ByVal Y As Long, Optional ByVal aFlag As EMonitorDefault = EMonitorDefault.ToNearest) As Monitor
    Set MonitorP = New Monitor: MonitorP.NewP X, Y, aFlag
End Function
Public Function MonitorR(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal aFlag As EMonitorDefault = EMonitorDefault.ToNearest) As Monitor
    Set MonitorR = New Monitor: MonitorR.NewR X, Y, Width, Height, aFlag
End Function
Public Function MonitorW(ByVal hwnd As LongPtr, Optional ByVal aFlag As EMonitorDefault = EMonitorDefault.ToNearest) As Monitor
    Set MonitorW = New Monitor: MonitorW.NewW hwnd, aFlag
End Function

Public Function MSCrypt(ByVal alg As ECryptHashAlgo) As MSCrypt
    Set MSCrypt = New MSCrypt: MSCrypt.New_ alg
End Function

Public Function MyThread(ith As IThreading, ByVal Index As Long, ByVal sIP As String) As MyThread
    Set MyThread = New MyThread: MyThread.New_ ith, Index, sIP
End Function

'Public Function MyThread(ith As IThreading, ByVal sIP As String) As MyThread
'    Set MyThread = New MyThread: MyThread.New_ ith, sIP
'End Function





















































Public Function SScreen(DstPB As PictureBox, SrcRect As WndRect) As SScreen
    Set SScreen = New SScreen: SScreen.New_ DstPB, SrcRect
End Function

Public Function Screenshot(PicBmp As StdPicBmp, SrcRect As WndRect) As Screenshot
    Set Screenshot = New Screenshot: Screenshot.New_ PicBmp, SrcRect
End Function

Public Function WinAPIRect(ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As WinAPIRect
    With WinAPIRect: .Left = X: .Top = Y: .Right = X + W: .Bottom = Y + H: End With
End Function
'
'Public Function WinAPIRect(ByVal L As Long, ByVal T As Long, ByVal R As Long, ByVal B As Long) As WinAPIRect
'    With WinAPIRect: .Left = L: .Top = T: .Right = R: .Bottom = B: End With
'End Function

Public Function WinAPIPoint(ByVal X As Long, ByVal Y As Long) As WinAPIPoint
    With WinAPIPoint: .X = X: .Y = Y: End With
End Function

Public Function WinAPISize(ByVal W As Long, ByVal H As Long) As WinAPISize
    With WinAPISize: .Width = W: .Height = H: End With
End Function

Public Function StdPicBmp(StdPic As StdPicture) As StdPicBmp
    Set StdPicBmp = New StdPicBmp: StdPicBmp.New_ StdPic
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

'Bsp: "DIN EN 10365", "HE-B 300", 300, 300, 11, 19, 27
'Public Function ProfilVollHI(ByVal Norm As String, ByVal Name As String, ByVal H_mm As Double, ByVal B_mm As Double, ByVal Tw_mm As Double, ByVal Tf_mm As Double, Optional ByVal R_mm As Double = 0) As ProfilVollHI
'    Set ProfilVollHI = New ProfilVollHI: ProfilVollHI.New_ Norm, Name, H_mm, B_mm, Tw_mm, Tf_mm, R_mm
'End Function

Public Function ProfilVollHIN(ByVal norm As String, ByVal Name As String, ByVal H_mm As Double, ByVal B_mm As Double, ByVal Tw_mm As Double, ByVal Tf_mm As Double, Optional ByVal R1_mm As Double = 0, Optional ByVal R2_mm As Double = 0, Optional ByVal Stg_Proz As Double = 0) As ProfilVollHIN
    Set ProfilVollHIN = New ProfilVollHIN: ProfilVollHIN.New_ norm, Name, H_mm, B_mm, Tw_mm, Tf_mm, R1_mm, R2_mm, Stg_Proz
End Function

'Bsp: "UPN 220", "DIN EN 10365", 220, 80, 9, , 12.5, 6.5
Public Function ProfilVollUN(ByVal norm As String, ByVal Name As String, ByVal H_mm As Double, ByVal B_mm As Double, ByVal Tw_mm As Double, ByVal Tf_mm As Double, Optional b1_mm, Optional ByVal R1_mm As Double = 0, Optional ByVal R2_mm As Double = 0, Optional Stg_Proz) As ProfilVollUN
    Set ProfilVollUN = New ProfilVollUN: ProfilVollUN.New_ norm, Name, H_mm, B_mm, Tw_mm, Tf_mm, b1_mm, R1_mm, R2_mm, Stg_Proz
End Function

Public Function ProfilVollRe(ByVal norm As String, ByVal Name As String, ByVal H_mm As Double, ByVal B_mm As Double) As ProfilVollRe
    Set ProfilVollRe = New ProfilVollRe: ProfilVollRe.New_ norm, Name, H_mm, B_mm
End Function

Public Function ProfilVollKr(ByVal D_mm As Double) As ProfilVollKr
    Set ProfilVollKr = New ProfilVollKr: ProfilVollKr.New_ D_mm
End Function

Public Function ProfilVollViertelKr(ByVal R_mm As Double, ByVal Drehung As EViertelKreisDrehung) As ProfilVollViertelKr
    Set ProfilVollViertelKr = New ProfilVollViertelKr: ProfilVollViertelKr.New_ R_mm, Drehung
End Function

Public Function ProfilVollViertelKrZwi(ByVal R_mm As Double, ByVal Drehung As EViertelKreisDrehung) As ProfilVollViertelKrZwi
    Set ProfilVollViertelKrZwi = New ProfilVollViertelKrZwi: ProfilVollViertelKrZwi.New_ R_mm, Drehung
End Function

'Hohlprofile
'Kreis & Ellipse
Public Function ProfilHohlKrEl(ByVal norm As String, ByVal Name As String, ByVal H_mm As Double, ByVal B_mm As Double, ByVal T_mm As Double, Optional ByVal IsKaltgefertigt As Boolean = False) As ProfilHohlKrEl
    Set ProfilHohlKrEl = New ProfilHohlKrEl: ProfilHohlKrEl.New_ norm, Name, H_mm, B_mm, T_mm, IsKaltgefertigt
End Function

'Quadrat & Rechteck
Public Function ProfilHohlQuRe(ByVal norm As String, ByVal Name As String, ByVal H_mm As Double, ByVal B_mm As Double, ByVal T_mm As Double, Optional ra, Optional Ri, Optional ByVal IsKaltgefertigt As Boolean = False) As ProfilHohlQuRe
    Set ProfilHohlQuRe = New ProfilHohlQuRe: ProfilHohlQuRe.New_ norm, Name, H_mm, B_mm, T_mm, ra, Ri, IsKaltgefertigt
End Function
Public Function Person(ByVal BirthDay As Date, ByVal City As City, ByVal Index As Long, ByVal Name As String) As Person
    Set Person = New Person: Person.New_ BirthDay, City, Index, Name
End Function

Public Function PersonDefault(City As City) As Person
    Set PersonDefault = MNew.Person(DateSerial(1980, 1, 1), City, 1, "Max Mustermann")
End Function

'Public Function DBcrud(Col As Collection, ListBox As ListBox, _
'                       BtnAdd As CommandButton, Optional BtnAddClone, Optional BtnInsert, Optional BtnInsertClone, Optional BtnEdit, _
'                       Optional BtnDelete, Optional BtnMoveUp, Optional BtnMoveDown, Optional BtnSortUp, Optional BtnSortDown, Optional BtnSearch) As DBcrud
'    Set DBcrud = New DBcrud: DBcrud.New_ Col, ListBox, BtnAdd, BtnAddClone, BtnInsert, BtnInsertClone, BtnEdit, BtnDelete, BtnMoveUp, BtnMoveDown, BtnSortUp, BtnSortDown, BtnSearch
'End Function
Public Function Point3D(ByVal TagName As String, ByVal X As Double, ByVal Y As Double, ByVal z As Double) As Point3D
    Set Point3D = New Point3D: Point3D.New_ TagName, X, Y, z
End Function


Public Function UUID() As GUID
    Set UUID = New GUID: UUID.Parse UUID.UUID
End Function

Public Function sscale(ScalableControl) As sscale
    Set sscale = New sscale: sscale.New_ ScalableControl
End Function

Public Function Twig(aName As String) As Twig
    Set Twig = New Twig: Twig.New_ aName
End Function

Public Function TreeListBox(aLB As ListBox, aTabSize As Long, aRoot As Twig) As TreeListBox
    Set TreeListBox = New TreeListBox: TreeListBox.New_ aLB, aTabSize, aRoot
End Function

Public Function Splitter(BolMDI As Boolean, Owner As Object, Container As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ BolMDI, Owner, Container, Name, LeftTop, RghtBot
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function VBPrjGroupN(Name As String) As VBPrjGroup
    Set VBPrjGroupN = New VBPrjGroup: VBPrjGroupN.NewN Name
End Function
Public Function VBPrjGroup(cid As ConfigIniDocument) As VBPrjGroup
    Set VBPrjGroup = New VBPrjGroup: VBPrjGroup.New_ cid
End Function

Public Function VBProject(cid As ConfigIniDocument) As VBProject
    Set VBProject = New VBProject: VBProject.New_ cid
End Function

Public Function VBProjProps(cid As ConfigIniDocument) As VBProjProps
    Set VBProjProps = New VBProjProps: VBProjProps.New_ cid
End Function
Public Function VBCodeFile(pfn As PathFileName, ByVal Name As String) As VBCodeFile
    Set VBCodeFile = New VBCodeFile: VBCodeFile.New_ pfn, Name
End Function

Public Function Version(ByVal aMajor As Long, ByVal aMinor As Long, Optional ByVal aBuild As Long = -1, Optional ByVal aRevision As Long = -1) As Version
    Set Version = New Version: Version.New_ aMajor, aMinor, aBuild, aRevision
End Function
'Public Function VersionS(ByVal VersionString As String) As Version
'    Set VersionS = New Version: VersionS.NewS VersionString
'End Function
'Public Function VersionA() As Version
'    Set VersionA = New Version: VersionA.NewA
'End Function
'Public Function VersionD(Optional DateValue) As Version
'    Set VersionD = New Version: VersionD.NewD DateValue
'End Function

Public Function Nullable(ByVal vt As EVbVarType) As Nullable
    Set Nullable = New Nullable: Nullable.New_ vt
End Function


Public Function VBClass(ByVal Modifier As EAcsModifier, ByVal Name As String, _
                        Optional ByVal HasCtorPair As Boolean = True, Optional ByVal HasClone As Boolean = True, Optional ByVal HasSetParam As Boolean = True, _
                        Optional ByVal HasToStr As Boolean = True, Optional ByVal HasKey As Boolean = False, Optional ByVal HasPtr As Boolean = False, _
                        Optional ByVal HasIsEqual As Boolean = False, Optional ByVal HasCompare As Boolean = False, Optional ByVal HasIsSame As Boolean = False, _
                        Optional ByVal BaseClass As VBClass = Nothing, Optional ByVal Interfaces As CCollection = Nothing) As VBClass
    Set VBClass = New VBClass: VBClass.New_ Modifier, Name, HasCtorPair, HasClone, HasSetParam, HasToStr, HasKey, HasPtr, HasIsEqual, HasCompare, HasIsSame, BaseClass, Interfaces
End Function

Public Function VBEnum(ByVal Modifier As EAcsModifier, ByVal Name As String, ByVal IsBitRow As Boolean) As VBEnum
    Set VBEnum = New VBEnum: VBEnum.New_ Modifier, Name, IsBitRow
End Function

Public Function VBEnumConst(ByVal Name As String, ByVal Value As Variant, ByVal IsHex As Boolean) As VBEnumConst
    Set VBEnumConst = New VBEnumConst: VBEnumConst.New_ Name, Value, IsHex
End Function

Public Function VBFunc(ByVal Modifier As EAcsModifier, ByVal Name As String, ByVal ReturnType As VbVarType, ByVal TypeName As String) As VBFunc
    Set VBFunc = New VBFunc: VBFunc.New_ Modifier, Name, ReturnType, TypeName
End Function

Public Function VBParam(ByVal Modifier As EArgModifier, ByVal Name As String, ByVal ReturnType As VbVarType, ByVal TypeName As String) As VBParam
    Set VBParam = New VBParam: VBParam.New_ Modifier, Name, ReturnType, TypeName
End Function

Public Function VBProj(ByVal Name As String) As VBProj
    Set VBProj = New VBProj: VBProj.New_ Name
End Function

Public Function VBProp(ByVal Modifier As EAcsModifier, ByVal Name As String, ByVal DataType As VbVarType, ByVal TypeName As String, Optional ByVal IsReadOnly As Boolean = True, Optional IsIndexed As Boolean = False, Optional IsCondensed As Boolean = False) As VBProp
    Set VBProp = New VBProp: VBProp.New_ Modifier, Name, DataType, TypeName, IsReadOnly, IsIndexed
End Function

Public Function VBSub(ByVal Modifier As EAcsModifier, ByVal Name As String, Parameters As CCollection) As VBSub
    Set VBFunc = New VBFunc: VBFunc.New_ Modifier, Name, Parameters
End Function

Public Function VBType(ByVal Modifier As EAcsModifier, ByVal Name As String, ByVal HasCtor As Boolean, ByVal HasToStr As Boolean) As VBType
    Set VBType = New VBType: VBType.New_ Modifier, Name, HasCtor, HasToStr
End Function

Public Function VBTypeVar(ByVal Name As String, ByVal DataType As VbVarType, Optional ByVal TypeName As String = "") As VBTypeVar
    Set VBTypeVar = New VBTypeVar: VBTypeVar.New_ Name, DataType, TypeName
End Function

Public Function Stahl(ByVal BaustahlArt As EBaustahlArt, ByVal fykMax As Double, ByVal GueteGruppeJRJ0 As String, ByVal Stahlgruppe As Integer, ByVal Zaehler As Integer, ByVal t_min As Double, ByVal t_max As Double, ByVal fyk As Double, ByVal fuk_min As Double, ByVal fuk_max As Double) As Stahl ', Optional ByVal LegierungsElemente As String = "")
    Set Stahl = New Stahl: Stahl.New_ BaustahlArt, fykMax, GueteGruppeJRJ0, Stahlgruppe, Zaehler, t_min, t_max, fyk, fuk_min, fuk_max ', LegierungsElemente
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function Variable(ByVal Name As String, ByVal Value As String) As Variable
    Set Variable = New Variable: Variable.New_ Name, Value
End Function

Public Function VariableR(inistr As String) As Variable
    Dim var As Variable: Set var = New Variable: If var.Read(inistr) Then Set VariableR = var
End Function


Public Function WndPicker(aTimer As Timer, aButton As CommandButton, ByVal bDrawFocusRect As Boolean) As WndPicker
    Set WndPicker = New WndPicker: WndPicker.New_ aTimer, aButton, bDrawFocusRect
End Function

Public Function WndInputs(ByVal hWndSender As LongPtr, ByVal hWndReceiver As LongPtr, ByVal Name As String, ByVal delaytime_ms As Integer) As WndInputs
    Set WndInputs = New WndInputs: WndInputs.New_ Name, hWndSender, hWndReceiver, delaytime_ms
End Function

Public Function WndInputMouse(ByVal Dx As Long, ByVal Dy As Long, ByVal MouseData As Long, ByVal flags As Long) As WndInputMouse ', ByVal aTime As Long
    Set WndInputMouse = New WndInputMouse: WndInputMouse.New_ Dx, Dy, MouseData, flags ', aTime
End Function

Public Function WndInputKeybd(ByVal VirtKey As EVirtualKeyCodes, ByVal Scan As Integer, ByVal flags As EKeyEventFlags) As WndInputKeybd ', ByVal aTime As Long
    Set WndInputKeybd = New WndInputKeybd: WndInputKeybd.New_ CInt(VirtKey), Scan, flags ', aTime
End Function

Public Function WndInputHardw(ByVal aMessage As Long, ByVal WParamL As Integer, ByVal WParamH As Integer) As WndInputHardw
    Set WndInputHardw = New WndInputHardw: WndInputHardw.New_ aMessage, WParamL, WParamH
End Function

Public Function WndInputDelay(ByVal Name As String, ByVal Milliseconds As Long) As WndInputDelay
    Set WndInputDelay = New WndInputDelay: WndInputDelay.New_ Name, Milliseconds
End Function

Public Function WndInputText(ByVal Name As String, ByVal Text As String, ByVal hWndSender As LongPtr, ByVal hWndReceiver As LongPtr) As WndInputText
    Set WndInputText = New WndInputText: WndInputText.New_ Name, Text, hWndSender, hWndReceiver
End Function

Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

Public Function NamedIBAN(ByVal aName As String, aIBAN As IBAN) As NamedIBAN
    Set NamedIBAN = New NamedIBAN: NamedIBAN.New_ aName, aIBAN
End Function

Public Function PictureBoxZoom(Window As Form, Canvas As PictureBox, aImage As StdPicture) As PictureBoxZoom
    Set PictureBoxZoom = New PictureBoxZoom: PictureBoxZoom.New_ Window, Canvas, aImage
End Function

Public Function ScannerTwain(Owner As Form) As ScannerTwain
    Set ScannerTwain = New ScannerTwain: ScannerTwain.New_ Owner
End Function


Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function


Public Function Nullable(ByVal vt As EVbVarType) As Nullable
    Set Nullable = New Nullable: Nullable.New_ vt
End Function




Public Function PersonalEvent(aDate As Date, aName As String, EventName As String) As PersonalEvent
    Set PersonalEvent = New PersonalEvent: PersonalEvent.New_ aDate, aName, EventName
End Function


Public Function PictureBoxZoom(Window As Form, Canvas As PictureBox, aImage As StdPicture) As PictureBoxZoom
    Set PictureBoxZoom = New PictureBoxZoom: PictureBoxZoom.New_ Window, Canvas, aImage
End Function

Public Function ScannerTwain(Owner As Form) As ScannerTwain
    Set ScannerTwain = New ScannerTwain: ScannerTwain.New_ Owner
End Function


Public Function Thread(ByVal aPriority As EThreadPriority, ByVal aAffinityMask As Long) As Thread
    Set Thread = New Thread: Thread.New_ aPriority, aAffinityMask
End Function

Public Function Thread(ByVal aPriority As EThreadPriority, ByVal aAffinityMask As Long) As Thread
    Set Thread = New Thread: Thread.New_ aPriority, aAffinityMask
End Function

Public Function Pipe(Optional ByVal Name As String, Optional bIsLocal As Boolean = False, _
                     Optional ByVal PipeAccess As EPipeAccess = PipeAccess_Duplex, Optional ByVal PipeType As EPipeType = PipeType_BYTE, _
                     Optional ByVal nInstances As Long = 255) As Pipe
    Set Pipe = New Pipe: Pipe.New_ Name, bIsLocal, PipeAccess, PipeType, nInstances
End Function


Public Function Range(ByVal MinValue As Double, ByVal MaxValue As Double, ByVal RangeType As ERangeType) As Range
    Set Range = New Range: Range.New_ MinValue, MaxValue, RangeType
End Function

Public Function RangeExMinExMax(ByVal MinValue As Double, ByVal MaxValue As Double) As RangeExMinExMax
    Set RangeExMinExMax = New RangeExMinExMax: RangeExMinExMax.New_ MinValue, MaxValue
End Function
Public Function RangeExMinInMax(ByVal MinValue As Double, ByVal MaxValue As Double) As RangeExMinInMax
    Set RangeExMinInMax = New RangeExMinInMax: RangeExMinInMax.New_ MinValue, MaxValue
End Function
Public Function RangeInMinExMax(ByVal MinValue As Double, ByVal MaxValue As Double) As RangeInMinExMax
    Set RangeInMinExMax = New RangeInMinExMax: RangeInMinExMax.New_ MinValue, MaxValue
End Function
Public Function RangeInMinInMax(ByVal MinValue As Double, ByVal MaxValue As Double) As RangeInMinInMax
    Set RangeInMinInMax = New RangeInMinInMax: RangeInMinInMax.New_ MinValue, MaxValue
End Function
    
Public Function Splitter(bMDI As Boolean, MyOwner As Object, MyContainer As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ bMDI, MyOwner, MyContainer, Name, LeftTop, RghtBot
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function Version(ByVal aMajor As Long, ByVal aMinor As Long, Optional ByVal aBuild As Long = -1, Optional ByVal aRevision As Long = -1) As Version
    Set Version = New Version: Version.New_ aMajor, aMinor, aBuild, aRevision
End Function
Public Function VersionS(ByVal VersionString As String) As Version
    Set VersionS = New Version: VersionS.NewS VersionString
End Function
Public Function VersionA() As Version
    Set VersionA = New Version: VersionA.NewA
End Function
Public Function VersionD(Optional DateValue) As Version
    Set VersionD = New Version: VersionD.NewD DateValue
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function StepDocument(aHeader As StepHeader, aData As StepData) As StepDocument
    Set StepDocument = New StepDocument: StepDocument.New_ aHeader, aData
End Function

Public Function StepHeader(ByVal FDescr As String, ByVal FImpl As String, ByVal PFNam As String, DateTimeStamp, ByVal Auth As String, ByVal Organis As String, ByVal FSchema As String) As StepHeader
    Set StepHeader = New StepHeader: StepHeader.New_ FDescr, FImpl, PFNam, DateTimeStamp, Auth, Organis, FSchema
End Function

Public Function StepObject(ByVal aHash As Long, ByVal aClassName As String) As StepObject
    Set StepObject = New StepObject: StepObject.New_ aHash, aClassName
End Function

'Public Function StepReader(aPathFileName As PathFileName) As StepReader
'    Set StepReader = New StepReader: StepReader.New_ aPathFileName
'End Function
Public Function StepReader(aStr As StreamStr) As StepReader
    Set StepReader = New StepReader: StepReader.New_ aStr
End Function

Public Function StepToken(aTokenType As EStepTokenType, aValue) As StepToken
    Set StepToken = New StepToken: StepToken.New_ aTokenType, aValue
End Function

Public Function StreamStr(ByVal aValue As String) As StreamStr
    Set StreamStr = New StreamStr: StreamStr.New_ aValue
End Function

Public Function StepTokenizer(aStream As StreamStr) As StepTokenizer
    Set StepTokenizer = New StepTokenizer: StepTokenizer.New_ aStream
End Function

Public Function nTupel(Of_Typ As VbVarType, ByVal nCount As Long) As nTupel
    Set nTupel = New nTupel: nTupel.New_ Of_Typ, nCount
End Function


Public Function nTupel(Of_Type As VbVarType, ByVal nCount As Long, ParamArray params()) As nTupel
    Set nTupel = New nTupel: nTupel.New_ Of_Type, nCount
    Dim i As Long, u As Long: u = UBound(params)
    If Of_Type = vbObject Then
        For i = 0 To u: Set nTupel.Values(i) = params(i): Next
    Else
        For i = 0 To u:     nTupel.Values(i) = params(i): Next
    End If
End Function




Public Function UTM32(ByVal Northing As Double, ByVal Easting As Double, ByVal UTMZone As String) As UTM32
    Set UTM32 = New UTM32: UTM32.New_ Northing, Easting, UTMZone
End Function
Public Function UTM32G(aGeoPos As GeoPos, elli As Ellipsoid) As UTM32
    Set UTM32G = New UTM32: UTM32G.NewG aGeoPos, elli
End Function

Public Function Func1(Obj As Object, ByVal FuncName As String) As Func1
    Set Func1 = New Func1: Func1.New_ Obj, FuncName
End Function
Public Function PropLet(Obj As Object, ByVal PropName As String) As PropLet
    Set PropLet = New PropLet: PropLet.New_ Obj, PropName
End Function


Public Function VBEnum(ByVal AccessModifier As VBAccessModifier, ByVal Name As String) As VBEnum
    Set VBEnum = New VBEnum: VBEnum.New_ AccessModifier, Name
End Function

Public Function VBConst(aEnum As VBEnum, ByVal AccessModifier As VBAccessModifier, ByVal Name As String, ByVal vType As VbVarType, Value As VBExprOrLiteral, ByVal DescriptionComment As String) As VBConst
    Set VBConst = New VBConst: VBConst.New_ aEnum, AccessModifier, Name, vType, Value, DescriptionComment
End Function

Public Function VBExprOrLiteral(ByVal Value, Optional ByVal vtid As Long = 0, Optional ByVal IsExpression As Boolean = False) As VBExprOrLiteral
    Set VBExprOrLiteral = New VBExprOrLiteral: VBExprOrLiteral.New_ Value, vtid, IsExpression
End Function

'
'Public Function DisplayDevices(ByVal DevName As String) As DisplayDevices
'    Set DisplayDevices = New DisplayDevices: DisplayDevices.New_ DevName
'End Function
'
'Public Function DeviceMode(ByVal DeviceName As String) As DeviceMode
'    Set DeviceMode = New DeviceMode: DeviceMode.New_ DeviceName
'End Function


'
'Public Function DisplayDevices(ByVal DevName As String) As DisplayDevices
'    Set DisplayDevices = New DisplayDevices: DisplayDevices.New_ DevName
'End Function
'
'Public Function DeviceMode(ByVal DeviceName As String) As DeviceMode
'    Set DeviceMode = New DeviceMode: DeviceMode.New_ DeviceName
'End Function

Public Function Window(ByVal Name As String, Optional ByVal Caption As String, Optional ByVal Style As EWndStyle = EWndStyle.VBFormStyle_Sizable, Optional ByVal StyleEx As EWndStyleEx = EWndStyleEx.VBFormStyleEx_Sizable) As Window
    Set Window = New Window: Window.New_ Name, Caption, Style, StyleEx
End Function


Public Function ResourceFile(aPFN As String) As ResourceFile
    Set ResourceFile = New ResourceFile: ResourceFile.New_ aPFN
End Function

Public Function Resource(ByVal HRSRC As LongPtr) As Resource
    Set Resource = New Resource: Resource.New_ HRSRC
End Function



Public Function PropGet(Obj As Object, ByVal PropName As String) As PropGet
    Set PropGet = New PropGet: PropGet.New_ Obj, PropName
End Function
Public Function PropGetObj(Obj As Object, ByVal PropName As String) As PropGetObj
    Set PropGetObj = New PropGetObj: PropGetObj.New_ Obj, PropName
End Function

Public Function PropLet(Obj As Object, ByVal PropName As String) As PropLet
    Set PropLet = New PropLet: PropLet.New_ Obj, PropName
End Function
Public Function PropSet(Obj As Object, ByVal PropName As String) As PropSet
    Set PropSet = New PropSet: PropSet.New_ Obj, PropName
End Function


Public Function Version(ByVal Major As Long, ByVal Minor As Long, Optional ByVal Build, Optional ByVal Revision) As Version
    Set Version = New Version: Version.New_ Major, Minor, Build, Revision
End Function

Public Function VersionS(ByVal Value As String) As Version
    Set VersionS = New Version: VersionS.Parse Value
End Function

Public Function OperatingSystem(Platform As PlatformID, aVersion As Version) As OperatingSystem
    Set OperatingSystem = New OperatingSystem: OperatingSystem.New_ Platform, aVersion
End Function

Public Function TempPathFileName(aFileName As String) As PathFileName
    Set TempPathFileName = MNew.PathFileName(PathFileNameStatic.TempPath, aFileName, ".tmp")
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function Splitter(BolMDI As Boolean, Owner As Object, Container As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ BolMDI, Owner, Container, Name, LeftTop, RghtBot
End Function

Public Function VBAProj(vbap As VBProject) As VBAProj
    Set VBAProj = New VBAProj: VBAProj.New_ vbap
End Function

Public Function VBARef(ByVal AReference As VBIDE.Reference) As VBARef
    Set VBARef = New VBARef: VBARef.New_ AReference
End Function

Public Function VBAForm(VBComp As VBComponent, Content As String, aFormFrx As VBAfrx) As VBAForm
    Set VBAForm = New VBAForm: VBAForm.New_ VBComp, Content, aFormFrx
End Function

Public Function VBAfrx(VBComp As VBComponent, Bytes() As Byte) As VBAfrx
    Set VBAfrx = New VBAfrx: VBAfrx.New_ VBComp, Bytes()
End Function

Public Function VBAModule(VBComp As VBComponent, Content As String) As VBAModule
    Set VBAModule = New VBAModule: VBAModule.New_ VBComp, Content
End Function

Public Function VBAClass(VBComp As VBComponent, Content As String) As VBAClass
    Set VBAClass = New VBAClass: VBAClass.New_ VBComp, Content
End Function


Public Function Splitter(BolMDI As Boolean, Owner As Object, Container As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ BolMDI, Owner, Container, Name, LeftTop, RghtBot
End Function

Public Function Nullable(ByVal vt As EVbVarType) As Nullable
    Set Nullable = New Nullable: Nullable.New_ vt
End Function

Public Function VVariant(aValue) As VVariant
    'Create a VVariant-object with a Variant
    Set VVariant = New VVariant: VVariant.New_ aValue
End Function

Public Function VVariantVt(vt As EVbVarType, aValue) As VVariant
    'Create a VVariant-object with a Variant and set the vartype yourself,
    'like e.g. give a signed Long and set vt to unsigned Long
    Set VVariantVt = New VVariant: VVariantVt.NewVt vt, aValue
End Function

Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

Public Function PFNDirectory(aRoot As PFNDirectory, Name As String) As PFNDirectory
    Set PFNDirectory = New PFNDirectory: PFNDirectory.New_ aRoot, Name
End Function

Public Function PFNDirectoryP(pfn As PathFileName) As PFNDirectory
    If Not pfn.IsPath Then Exit Function
    Dim root As PFNDirectory: Set root = MNew.PFNDrive(pfn.Drive)
    Dim i As Long
    For i = 0 To pfn.PathCount - 1
        Set root = MNew.PFNDirectory(root, pfn.PathI(i))
    Next
    Set PFNDirectoryP = root
End Function

Public Function PFNDrive(ByVal aLetter As String) As PFNDrive
    Set PFNDrive = New PFNDrive: PFNDrive.New_ aLetter
End Function

Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

'Public Function PFNStreamW(PFN As PathFileName) As PFNStreamW
'    Set PFNStreamW = New PFNStreamW: PFNStreamW.New_ PFN
'End Function

Public Function Pdf(ByVal aPdfPFN As PathFileName) As Pdf
    Set Pdf = New Pdf: Pdf.New_ aPdfPFN
End Function

'Public Function ListOfPdf(aPDFTK As PDFToolKit) As ListOfPdf
'    Set ListOfPdf = New ListOfPdf: ListOfPdf.New_ aPDFTK
'End Function

Public Function PDFToolKit(PFN_PdfTK_exe As PathFileName) As PDFToolKit
    Set PDFToolKit = New PDFToolKit: PDFToolKit.New_ PFN_PdfTK_exe
End Function

Public Function Twig(aName As String) As Twig
    Set Twig = New Twig
    Twig.New_ aName
End Function

Public Function Splitter(BolMDI As Boolean, MyOwner As Object, MyContainer As Object, Name As String, LeftTop As Control, LeftBot As Control) As Splitter
    Set Splitter = New Splitter
    Splitter.New_ BolMDI, MyOwner, MyContainer, Name, LeftTop, LeftBot
End Function

Public Function TreeListBox(aLB As ListBox, aTabSize As Long, aRoot As Twig) As TreeListBox
    Set TreeListBox = New TreeListBox
    TreeListBox.New_ aLB, aTabSize, aRoot
End Function

Public Function ServerInfo() As ServerInfo
    Set ServerInfo = New ServerInfo
End Function

Public Function UserInfo() As UserInfo
    Set UserInfo = New UserInfo
End Function

Public Function ScrollableTextBox(Owner As Form, aTB As TextBox, ByVal Panel As PictureBox, ByVal MinValue As Double, ByVal MaxValue As Double, ByVal CountSteps As Long, StrFormat As String, Optional Btn As CommandButton, Optional ScrollComboBox As ComboBox = Nothing, Optional ByVal InverseScroll As Boolean = False) As ScrollableTextBox
    Set ScrollableTextBox = New ScrollableTextBox: ScrollableTextBox.New_ Owner, aTB, Panel, MinValue, MaxValue, CountSteps, StrFormat, Btn, ScrollComboBox, InverseScroll
End Function



Public Function VVariant(aValue) As VVariant
    'Create a VVariant-object with a Variant
    Set VVariant = New VVariant: VVariant.New_ aValue
End Function
Public Function VVariantVt(vt As EVbVarType, aValue) As VVariant
    'Create a VVariant-object with a Variant and set the vartype yourself,
    'like e.g. give a signed Long and set vt to unsigned Long
    Set VVariantVt = New VVariant: VVariantVt.NewVt vt, aValue
End Function

Public Function VVariantPtr(aValue) As VVariantPtr
    'Create a VVariantPtr-object with a Variant
    Set VVariantPtr = New VVariantPtr: VVariantPtr.New_ aValue
End Function
Public Function VVariantPtrVt(vt As EVbVarType, aValue) As VVariantPtr
    'Create a VVariant-object with a Variant and set the vartype yourself,
    'like e.g. give a signed Long and set vt to unsigned Long
    Set VVariantPtrVt = New VVariantPtr: VVariantPtrVt.NewVt vt, aValue
End Function



Public Function UUID() As GUID
    Set UUID = New GUID: UUID.Parse UUID.UUID
End Function

' ############################## '    PropertyStore    ' ############################## '
Public Function PropertyStore(ByVal FileName As String) As PropertyStore
    Set PropertyStore = New PropertyStore: PropertyStore.New_ FileName
End Function

' ############################## '    PropKey    ' ############################## '
Public Function PropKey(aFMTID_Guid As String, aPID As Long, Optional aName As String) As PropKey
    Set PropKey = New PropKey: PropKey.New_ aFMTID_Guid, aPID, aName
End Function

Public Function PropKeyPkhe(aPkhe As PropKeyHEntry) As PropKey
    Set PropKeyPkhe = MNew.PropKey(aPkhe.FmtGuid, CLng(aPkhe.PIDValue), aPkhe.Name)
End Function

' ############################## '    PathFileName    ' ############################## '
Public Function PathFileName(ByVal aPathFileName As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function


'Public Function ToPropKey() As PropKey
'    Dim PID As Long: PID = CLng(PIDValue)
'    Set ToPropKey = MNew.PropKey(FmtGuid, PID, Name)
'End Function

'PropKeyHEntry:
'Public Function ToVVariant() As VVariant
'    Dim vt As EVbVarType: vt = MEVbVarType.EVbVarType_Parse(Me.PKVarTyp)
'    Set ToVVariant = MNew.VVariantVt(vt)
'End Function
'
'Public Function ToPossFileProperty() As FileProperty
'    Set ToPossFileProperty = MNew.FileProperty(Me.ToPropKey, Me.ToVVariant)
'End Function

Public Function VVariant(aValue) As VVariant
    'Create a VVariant-object with a Variant
    Set VVariant = New VVariant: VVariant.New_ aValue
End Function
Public Function VVariantVt(Optional ByVal vt As EVbVarType = EVbVarType.vbIllegal, Optional ByVal aValue = Empty) As VVariant
    'Create a VVariant-object with a Variant and set the vartype yourself,
    'like e.g. give a signed Long and set vt to unsigned Long
    Set VVariantVt = New VVariant: VVariantVt.NewVt vt, aValue
End Function

' ############################## '    PropKey    ' ############################## '
Public Function PropKey(aFMTID_Guid As String, aPID As Long, Optional aName As String) As PropKey
    Set PropKey = New PropKey: PropKey.New_ aFMTID_Guid, aPID, aName
End Function

Public Function PropKeyPkhe(aPkhe As PropKeyHEntry) As PropKey
    Set PropKeyPkhe = MNew.PropKey(aPkhe.FmtGuid, CLng(aPkhe.PIDValue), aPkhe.Name)
End Function


Public Function PathFileName(ByVal aPathFileName As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function UUID() As GUID
    Set UUID = New GUID: UUID.Parse UUID.UUID
End Function



Public Function PropertyStore(ByVal FileName As String) As PropertyStore
    Set PropertyStore = New PropertyStore: PropertyStore.New_ FileName
End Function

Public Function PropKey(aFMTID_Guid As String, aPID As Long, Optional aName As String) As PropKey
    Set PropKey = New PropKey: PropKey.New_ aFMTID_Guid, aPID, aName
End Function




Public Function UUID() As GUID
    Set UUID = New GUID: UUID.Parse UUID.UUID
End Function

Public Function PropertyStore(ByVal FileName As String) As PropertyStore
    Set PropertyStore = New PropertyStore: PropertyStore.New_ FileName
End Function

Public Function PropKey(aFMTID_Guid As String, aPID As Long, Optional aName As String) As PropKey
    Set PropKey = New PropKey: PropKey.New_ aFMTID_Guid, aPID, aName
End Function

'Jost Schwiders  StringBuilder www.VBTec.de 'originally it is called Concat
Public Function StringBuilderJS(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilderJS
    Set StringBuilderJS = New StringBuilderJS: StringBuilderJS.New_ Value, startIndex, Length, Capacity, maxCapacity
End Function

'Oliver Meyers  StringBuilder www.MBO-Ing.com
Public Function stringBuilder(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As stringBuilder
    Set stringBuilder = New stringBuilder: stringBuilder.New_ Value, startIndex, Length, Capacity, maxCapacity
End Function

'Steve McMahons StringBuilder www.VBAccelerator.com
Public Function StringBuilderSM(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilderSM
    Set StringBuilderSM = New StringBuilderSM: StringBuilderSM.ChunkSize = Capacity
End Function

'only a class around VB's original String concatenation with the "&"-operator
Public Function StringBuilderVB(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilderVB
    Set StringBuilderVB = New StringBuilderVB: StringBuilderVB.New_ Value, startIndex, Length, Capacity, maxCapacity
End Function

'##############################'   SBBench   '##############################'
Public Function SBBench(aFncPtr As Long, aSBBenchSituation As SBBenchSituation) As SBBench
    Set SBBench = New SBBench: SBBench.New_ aFncPtr, aSBBenchSituation
End Function

Public Function SBBenchSituation(Optional aXAvgRuns As Long = -1, Optional aStrVal As String = vbNullString, Optional aStrLen As Long = -1, Optional aNLoops As Long = -1) As SBBenchSituation
    Set SBBenchSituation = New SBBenchSituation: SBBenchSituation.New_ aXAvgRuns, aStrVal, aStrLen, aNLoops
End Function


Public Function PossFileProp(aName As String, aGUID As String, aPID As Long, aType As String) As PossFileProp
    Set PossFileProp = New PossFileProp: PossFileProp.New_ aName, aGUID, aPID, aType
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                         Optional ByVal aFileName As String, _
                         Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

'QuickSorter
Public Function QuickSorter(aCmp As List) As QuickSorter ' , aSwp As Object) As QuickSorter ' As IComparable, ISwappable
    Set QuickSorter = New QuickSorter: QuickSorter.New_ aCmp ', aSwp
End Function


Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

Public Function PFNDateTime(aPFN As PathFileName) As PFNDateTime
    Set PFNDateTime = New PFNDateTime: PFNDateTime.New_ aPFN
End Function

Public Function PropertyChangedEventArgs(ByVal aPropName As String) As PropertyChangedEventArgs
    Set PropertyChangedEventArgs = New PropertyChangedEventArgs: PropertyChangedEventArgs.New_ aPropName
End Function

Public Function ShellLink(aPFN As String) As ShellLink
    Set ShellLink = New ShellLink: ShellLink.New_ aPFN
End Function
'
Public Function PFNStream(aPath As String, Optional aFileName As String, Optional aExt As String, Optional FMode As FileMode = FileMode_Random, Optional FAccess As FileAccess = FileAccess_None, Optional FShare As FileShare = FileShare_None) As PFNStream
      Set PFNStream = New PFNStream: PFNStream.New_ aPath, aFileName, aExt, FMode, FAccess, FShare
End Function

Public Function PFNStreamP(aPFN As PathFileName, Optional FMode As FileMode = FileMode_Random, Optional FAccess As FileAccess = FileAccess_None, Optional FShare As FileShare = FileShare_None) As PFNStream
    Set PFNStreamP = New PFNStream: PFNStreamP.NewP aPFN, FMode, FAccess, FShare
End Function

Public Function New_HtmlElem(aName As String) As HtmlElem
    Set New_HtmlElem = New HtmlElem
    New_HtmlElem.New_ aName
End Function


Public Function ToolBar(aName As String, aForm As Form, aPBAsPanel As PictureBox) As ToolBar
    Set ToolBar = New ToolBar: ToolBar.New_ aName, aForm, aPBAsPanel
End Function

Public Function TBButton(aName As String, aOwner As Form, aPanel As PictureBox, aBtn As CommandButton, aPic As StdPicture) As TBButton
    Set TBButton = New TBButton: TBButton.New_ aName, aOwner, aPanel, aBtn, aPic
End Function


Public Function PEDocument(aPFN As String) As PEDocument
    Set PEDocument = New PEDocument: PEDocument.New_ aPFN
End Function
Public Function PEDirectoryEntry(aDataDir As PEDataDirectory, ByVal aIndex As Long) As PEDirectoryEntry
    Set PEDirectoryEntry = New PEDirectoryEntry: PEDirectoryEntry.New_ aDataDir, aIndex
End Function
    
Public Function WaveSound(wavformat As WaveFormat) As WaveSound
    Set WaveSound = New WaveSound: WaveSound.New_ wavformat
End Function

Public Function WaveView(aPBView As PictureBox) As WaveView
    Set WaveView = New WaveView: WaveView.New_ aPBView
End Function

Public Function WaveSound(ByVal bps As Integer, ByVal nChannels As Integer, ByVal sps As Long) As WaveSound
    Set WaveSound = New WaveSound: WaveSound.New_ bps, nChannels, sps
End Function
Public Function WaveSoundF(aFileName As String) As WaveSound
    Set WaveSoundF = New WaveSound: WaveSoundF.NewF aFileName
End Function

Public Function QrDecomposition(Value As Matrix) As QrDecomposition
    Set QrDecomposition = New QrDecomposition: QrDecomposition.New_ Value
End Function

Public Function SingularValueDecomposition(Value As Matrix) As SingularValueDecomposition
    Set SingularValueDecomposition = New SingularValueDecomposition: SingularValueDecomposition.New_ Value
End Function



Public Function Snapshot(ByVal flags As SnapShotFlags, Optional ByVal PID As Long) As Snapshot
    Set Snapshot = New Snapshot: Snapshot.New_ flags, PID
End Function

Public Function ProcessEntry(aSnapShot As Snapshot) As ProcessEntry
    Set ProcessEntry = New ProcessEntry: ProcessEntry.New_ aSnapShot
End Function

Public Function ThreadEntry(aSnapShot As Snapshot) As ThreadEntry
    Set ThreadEntry = New ThreadEntry: ThreadEntry.New_ aSnapShot
End Function


'##############################'    TabControl   '##############################'
Public Function TabControl(MyOwner As Form, MyContainer As PictureBox, MyName As String) As TabControl
    Set TabControl = New TabControl
    Call TabControl.New_(MyOwner, MyContainer, MyName)
End Function

Public Function Splitter(BolMDI As Boolean, Owner As Object, Container As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ BolMDI, Owner, Container, Name, LeftTop, RghtBot
End Function


Function StdFont(ByVal Name As String, ByVal Size As Single, _
                 Optional ByVal bBold As Boolean = False, _
                 Optional ByVal bItalic As Boolean = False, _
                 Optional ByVal aWeight As Long = 1, _
                 Optional ByVal aCharSet As Long = 1, _
                 Optional ByVal bUnderline As Boolean = False, _
                 Optional ByVal bStrikeThru As Boolean = False) As StdFont
    Set StdFont = New StdFont
    With StdFont
        .Name = Name
        .Size = Size
        .Bold = bBold
        .Italic = bItalic
        .Weight = aWeight
        .Charset = aCharSet
        .Underline = bUnderline
        .Strikethrough = bStrikeThru
    End With
End Function

Function StdFontC(ByVal Other As StdFont) As StdFont
    With Other
        Set StdFontC = MNew.StdFont(.Name, .Size, .Bold, .Italic, .Weight, .Charset, .Underline, .Strikethrough)
    End With
End Function

Public Sub StdFont_Assign(fntDst As StdFont, fntSrc As StdFont)
    With fntDst
        .Name = fntSrc.Name
        .Size = fntSrc.Size
        .Bold = fntSrc.Bold
        .Italic = fntSrc.Italic
        .Weight = fntSrc.Weight
        .Charset = fntSrc.Charset
        .Underline = fntSrc.Underline
        .Strikethrough = fntSrc.Strikethrough
    End With
End Sub

Public Function Splitter(BolMDI As Boolean, MyOwner As Object, MyContainer As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ BolMDI, MyOwner, MyContainer, Name, LeftTop, RghtBot
End Function


Public Function stringBuilder(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As stringBuilder
  Set stringBuilder = New stringBuilder: stringBuilder.New_ Value, startIndex, Length, Capacity, maxCapacity
End Function

Public Function StreamReader(ByVal aStream As Stream, Optional ByVal Encoding, Optional ByVal detectEncodingFromByteOrderMarks As Boolean, Optional ByVal bufferSize As Long) As StreamReader
  Set StreamReader = New StreamReader: StreamReader.New_ aStream, Encoding, detectEncodingFromByteOrderMarks, bufferSize
End Function

Public Function StreamReaderP(ByVal aPath As String, Optional ByVal Encoding, Optional ByVal detectEncodingFromByteOrderMarks As Boolean, Optional ByVal bufferSize As Long) As StreamReader
  Set StreamReaderP = New StreamReader: StreamReaderP.NewP aPath, Encoding, detectEncodingFromByteOrderMarks, bufferSize
End Function

Public Function TaggedImageFile(aFileName As String) As TaggedImageFile
    Set TaggedImageFile = New TaggedImageFile: TaggedImageFile.New_ aFileName
End Function

Public Function TaggedImageFile(aFileName As String) As TaggedImageFile
    Set TaggedImageFile = New TaggedImageFile: TaggedImageFile.New_ aFileName
End Function



Public Function Winkel(inRadians As Double) As Winkel
    MMain.InitMathConsts
    Set Winkel = New Winkel
    Winkel.New_ inRadians
End Function

Public Function WinkelDeg(inDegrees As Double) As Winkel
    MMain.InitMathConsts
    Set WinkelDeg = New Winkel
    WinkelDeg.NewDeg inDegrees
End Function

Public Function WinkelGon(InGon As Double) As Winkel
    MMain.InitMathConsts
    Set WinkelGon = New Winkel
    WinkelGon.NewGon InGon
End Function

Public Function WaveSound(ByVal WaveFmtTag As Integer, ByVal BitsPerSample As Integer, ByVal nChannels As Integer, ByVal SampleRate As Long) As WaveSound
    Set WaveSound = New WaveSound
    WaveSound.New_ WaveFmtTag, BitsPerSample, nChannels, SampleRate
End Function

Public Function WaveSoundFlt32(ByVal nChannels As Integer, ByVal SampleRate As Long) As WaveSound
    Set WaveSoundFlt32 = WaveSound(3, 32, nChannels, SampleRate)
End Function


Public Function Operator(aAmplitude As Operator, aOscillator As Oscillator, ByVal aFrequency As Double, aModulator As Operator) As Operator
    Set Operator = New Operator
    Operator.New_ aAmplitude, aOscillator, aFrequency, aModulator
End Function

Public Function OperatorConst(ByVal aValue As Double) As OperatorConst
    Set OperatorConst = New OperatorConst
    OperatorConst.New_ aValue
End Function

Public Function OperatorLinear(ByVal aStartValue As Double, ByVal aIncrement As Double) As OperatorLinear
    Set OperatorLinear = New OperatorLinear
    OperatorLinear.New_ aStartValue, aIncrement
End Function


Public Function WndRect(r As WinAPIRect) As WndRect
    Set WndRect = New WndRect: WndRect.New_ r
End Function
Public Function WndRectHWnd(ByVal hwnd As LongPtr) As WndRect
    Set WndRectHWnd = New WndRect: WndRectHWnd.NewFromHWnd hwnd
End Function
Public Function WndRectFromMousePoint(P As WinAPIPoint) As WndRect
    Set WndRectFromMousePoint = New WndRect: WndRectFromMousePoint.NewFromMousePoint P
End Function

Public Function XTimer(Listener As IListenXTimer, ByVal Interval_ms As Single) As XTimer
    Set XTimer = New XTimer: XTimer.New_ Listener, Interval_ms
End Function

Public Function XTimerL(Listener As IListenXTimer, ByVal Interval_ms As Single) As XTimerL
    Set XTimerL = New XTimerL: XTimerL.New_ Listener, Interval_ms
End Function

