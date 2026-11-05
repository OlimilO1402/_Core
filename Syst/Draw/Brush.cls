VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Brush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'lbStyle
Private Const BS_SOLID         As Long = 0 'Einfarbiger Pinsel.
Private Const BS_NULL          As Long = 1 'Identisch mit BS_HOLLOW.
Private Const BS_HOLLOW        As Long = 1 'Hohlpinsel.
Private Const BS_HATCHED       As Long = 2 'Geschlüpfter Pinsel.
Private Const BS_PATTERN       As Long = 3 'Musterpinsel, der durch eine Speicherbitbitte definiert wird.
Private Const BS_DIBPATTERN    As Long = 5 'Ein Musterpinsel, der durch eine DIB-Spezifikation (Device-Independent Bitmap) definiert wird. Wenn lbStyle BS_DIBPATTERN ist, enthält das lbHatch-Element einen Handle für einen gepackten DIB. Weitere Informationen finden Sie unter Diskussion in lbHatch.
Private Const BS_DIBPATTERNPT  As Long = 6 'Ein Musterpinsel, der durch eine DIB-Spezifikation (Device-Independent Bitmap) definiert wird. Wenn lbStyle BS_DIBPATTERNPT ist, enthält das lbHatch-Element einen Zeiger auf einen gepackten DIB. Weitere Informationen finden Sie unter Diskussion in lbHatch.
Private Const BS_PATTERN8X8    As Long = 7 'Siehe BS_PATTERN.
Private Const BS_DIBPATTERN8X8 As Long = 8 'Siehe BS_DIBPATTERN.

Public Enum EBrushStyle
    bsSolid = 0
    bsNull = 1
    'bsHollow = 1
    bsHatched = 2
    bsPattern = 3
    bsDIBPattern = 5
    bsDIBPatternPT = 6
    bsPattern8x8 = 7
    bsDIBPattern8x8 = 8
End Enum

'lbColor
'If lbStyle is BS_HATCHED or BS_SOLID, lbColor is a COLORREF color value
Private Const DIB_RGB_COLORS As Long = 0 'Die Farbtabelle enthält RGB-Literalwerte.
Private Const DIB_PAL_COLORS As Long = 1 'Die Farbtabelle besteht aus einem Array von 16-Bit-Indizes in der derzeit realisierten logischen Palette.

'lbHatch
'A hatch style. The meaning depends on the brush style defined by lbStyle.

'typedef struct tagLOGBRUSH {
'  UINT      lbStyle;
'  COLORREF  lbColor;
'  ULONG_PTR lbHatch;
'} LOGBRUSH, *PLOGBRUSH, *NPLOGBRUSH, *LPLOGBRUSH;

Private Type LOGBRUSH
    lbStyle As Long    ' UINT
    lbColor As Long    ' COLORREF
    lbHatch As LongPtr ' ULONG_PTR
End Type ' } LOGBRUSH, *PLOGBRUSH, *NPLOGBRUSH, *LPLOGBRUSH;


'CreateBrushIndirect     Creates a brush with a specified style, color, and pattern
Private Declare Function CreateBrushIndirect Lib "gdi32" (ByRef lpLogBrush As LOGBRUSH) As LongPtr
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hGDIObj As LongPtr) As LongPtr
Private Declare Function DeleteObject Lib "gdi32" (ByVal hGDIObj As LongPtr) As Long


'CreateDIBPatternBrushPt     Creates a brush with the pattern from a DIB
'Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (ByRef lpPackedDIB As Any, ByVal iUsage As Long) As LongPtr
'CreateHatchBrush    Creates a brush with a hatch pattern and color
'Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As LongPtr
'CreatePatternBrush  Creates a brush with a bitmap pattern
'Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As LongPtr) As LongPtr
'CreateSolidBrush    Creates a brush with a solid color
'Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr

'GetBrushOrgEx   Gets the brush origin for a device context
'Private Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hDC As LongPtr, ByRef lpPoint As POINTAPI) As Long
'SetBrushOrgEx   Sets the brush origin for a device context
'Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As LongPtr, ByVal nXOrg As Long, ByVal nYOrg As Long, ByRef lppt As POINTAPI) As Long

'GetSysColorBrush    Gets a handle to a brush that corresponds to a color index
'Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
'PatBlt  Paints a rectangle
'Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
'SetDCBrushColor     Sets the current device context brush color.
'Private Declare Function SetDCBrushColor Lib "gdi32" (ByVal hDC As LongPtr, ByVal colorref As Long) As Long

Private m_LogObj As LOGBRUSH
Private m_Handle As LongPtr
Private m_Hatch  As Bitmap

Private Sub Class_Initialize(): Me.New_:                End Sub
Private Sub Class_Terminate():  Me.Dispose:             End Sub

Friend Sub New_(Optional BrushStyle As EBrushStyle = EBrushStyle.bsSolid, Optional Color As Long = vbBlue, Optional Hatch As Bitmap = Nothing)
    With m_LogObj
        .lbStyle = BrushStyle
        .lbColor = Color
        If Not Hatch Is Nothing Then
            Set m_Hatch = Hatch
            .lbHatch = m_Hatch.Handle
        End If
    End With
    Create
End Sub

Friend Sub NewPtr(ByVal lpLogObj As LongPtr)
    RtlMoveMemory ByVal Me.Ptr, ByVal lpLogObj, LenB(m_LogObj)
    Create
End Sub

'Copy to MNew
'Public Function Brush(Optional ByVal BrushStyle As EBrushStyle = EBrushStyle.bsSolid, Optional ByVal Color As Long = vbBlue, Optional Hatch As Bitmap = Nothing) As Brush
'    Set Brush = New Brush: Brush.New_ BrushStyle, Color, Hatch
'End Function

Public Function NewC(other As Brush)
    'RtlMoveMemory m_LogObj, other.Ptr, LenB(m_LogObj)
    NewPtr other.Ptr
End Function

Public Sub Dispose():    If m_Handle Then DeleteObject m_Handle: m_Handle = 0
End Sub

Public Function Clone(Optional ByVal deep As Boolean = False) As Pen: Set Clone = New Pen: Clone.NewC Me: End Function

Public Sub SetParams(Optional Style, Optional Color, Optional Hatch)
    If Not IsMissing(Style) Then Me.Style = CLng(Style)
    If Not IsMissing(Color) Then Me.Color = CLng(Color)
    If Not IsMissing(Hatch) Then Set Me.Hatch = Hatch
End Sub

Public Function Create() As LongPtr: m_Handle = CreateBrushIndirect(m_LogObj): Create = m_Handle: End Function

Public Property Get Ptr() As LongPtr:                   Ptr = VarPtr(m_LogObj):        End Property

Public Property Get Handle() As LongPtr:                Handle = m_Handle:             End Property

Public Property Get Style() As EBrushStyle:            Style = m_LogObj.lbStyle:       End Property
Friend Property Let Style(ByVal Value As EBrushStyle): m_LogObj.lbStyle = Value:       End Property

Public Property Get Color() As Long:                   Color = m_LogObj.lbColor:       End Property
Public Property Let Color(ByVal Value As Long):        m_LogObj.lbColor = Value:       End Property

Public Property Get Hatch() As Bitmap:                  Set Hatch = m_Hatch:           End Property
Friend Property Set Hatch(ByVal Value As Bitmap):       Set m_Hatch = Value:           End Property

' v ' ############################## ' v '    all Enums    ' v ' ############################## ' v '
' v ' ############################## ' v '    EPenStyle    ' v ' ############################## ' v '
Public Function EBrushStyle_ToStr(ByVal e As EBrushStyle) As String
    Dim s As String
    Select Case e
    Case EBrushStyle.bsSolid:         s = "Solid" '& vbTab & "______"
    Case EBrushStyle.bsNull:          s = "Null" '& vbTab & "------"
    Case EBrushStyle.bsHatched:       s = "Hatched" '& vbTab & "......"
    Case EBrushStyle.bsPattern:       s = "Pattern" '& vbTab & "-.-.-."
    Case EBrushStyle.bsDIBPattern:    s = "DIBPattern" '& vbTab & "-..-.."
    Case EBrushStyle.bsDIBPatternPT:  s = "DIBPatternPT" '& vbTab & "      "
    Case EBrushStyle.bsPattern8x8:    s = "Pattern8x8" '& vbTab & "_"
    Case EBrushStyle.bsDIBPattern8x8: s = "DIBPattern8x8" '& vbTab & "_"
    End Select
    EBrushStyle_ToStr = s
End Function

Public Function EBrushStyle_Parse(ByVal s As String) As EBrushStyle
    s = Trim(Split(s, vbTab)(0))
    Dim e As EBrushStyle
    Select Case s
    Case "Solid":         e = EBrushStyle.bsSolid
    Case "Null":          e = EBrushStyle.bsNull
    Case "Hatched":       e = EBrushStyle.bsHatched
    Case "Pattern":       e = EBrushStyle.bsPattern
    Case "DIBPattern":    e = EBrushStyle.bsDIBPattern
    Case "DIBPatternPT":  e = EBrushStyle.bsDIBPatternPT
    Case "Pattern8x8":    e = EBrushStyle.bsPattern8x8
    Case "DIBPattern8x8": e = EBrushStyle.bsDIBPattern8x8
    End Select
    EBrushStyle_Parse = e
End Function

Public Sub EBrushStyle_ToCBLB(ByVal CBLB As ComboBox)
    CBLB.Clear: Dim i As Long, s As String: For i = 0 To 8: s = EBrushStyle_ToStr(i): If Len(s) Then CBLB.AddItem s
    Next
End Sub

