VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'lopnStyle
Private Const PS_SOLID       As Long = 0 ' The pen is solid.
Private Const PS_DASH        As Long = 1 ' The pen is dashed.
Private Const PS_DOT         As Long = 2 ' The pen is dotted.
Private Const PS_DASHDOT     As Long = 3 ' The pen has alternating dashes and dots.
Private Const PS_DASHDOTDOT  As Long = 4 ' The pen has dashes and double dots.
Private Const PS_NULL        As Long = 5 ' The pen is invisible.
Private Const PS_INSIDEFRAME As Long = 6 ' The pen is solid. When this pen is used in any GDI drawing function that takes a bounding rectangle, the dimensions of the figure are shrunk so that it fits entirely in the bounding rectangle, taking into account the width of the pen. This applies only to geometric pens.

Private Const PS_USERSTYLE         As Long = 7
Private Const PS_ALTERNATE         As Long = 8
Private Const PS_MAXLINKTYPES      As Long = 8
Private Const PS_ONLY_PRECIS       As Long = 10

Private Const PS_STYLE_MASK        As Long = &HF&

Private Const PS_JOIN_ROUND        As Long = &H0&
Private Const PS_COSMETIC          As Long = &H0&
Private Const PS_ENDCAP_ROUND      As Long = &H0&

Private Const PS_ENDCAP_SQUARE     As Long = &H100&
Private Const PS_ENDCAP_FLAT       As Long = &H200&
Private Const PS_ENDCAP_MASK       As Long = &HF00&

Private Const PS_JOIN_BEVEL        As Long = &H1000&
Private Const PS_JOIN_MITER        As Long = &H2000&
Private Const PS_JOIN_MASK         As Long = &HF000&
Private Const PS_GEOMETRIC         As Long = &H10000
Private Const PS_OPENTYPE_FONTTYPE As Long = &H10000
Private Const PS_TYPE_MASK         As Long = &HF0000

'Private Const PS_BASIC_DIRECTORY As String = "\x02\x55\x4\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"
'Private Const PS_SECURITY_DIRECTORY As String = "\x02\x55\x4\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"
'Private Const PS_EXCHANGE_DIRECTORY As String = "\x08\x2A\x86\x48\x86\xF7\x14\x01\x02\x00\x00\x00\x00\x00\x00\x00"

Public Enum EPenStyle
    psSolid = 0
    psDash = 1
    psDot = 2
    psDashDot = 3
    psDashDotDot = 4
    psNull = 5
    psInsideFrame = 6
End Enum


'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-logpen
'typedef struct tagLOGPEN {
'  UINT     lopnStyle;
'  POINT    lopnWidth;
'  COLORREF lopnColor;
'} LOGPEN, *PLOGPEN, *NPLOGPEN, *LPLOGPEN;

Private Type LOGPEN
    lopnStyle  As Long
    lopnWidthX As Long
    lopnWidthY As Long
    lopnColor  As Long '0x00bbggrr
End Type
'const COLORREF rgbRed   =  0x000000FF;
'const COLORREF rgbGreen =  0x0000FF00;
'const COLORREF rgbBlue  =  0x00FF0000;
'const COLORREF rgbBlack =  0x00000000;
'const COLORREF rgbWhite =  0x00FFFFFF;
'Example from Windows Classic Samples on GitHub.
'https://github.com/microsoft/Windows-classic-samples

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createpen
Private Declare Function CreatePen Lib "gdi32" (ByVal iStyle As Long, ByVal cWidth As Long, ByVal crColor As Long) As Long
'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createpenindirect
Private Declare Function CreatePenIndirect Lib "gdi32" (ByRef lpLogPen As LOGPEN) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long

Private m_LogObj As LOGPEN
Private m_Handle As LongPtr

Private Sub Class_Initialize(): Me.New_:                End Sub
Private Sub Class_Terminate():  Me.Dispose:             End Sub

Friend Sub New_(Optional PenStyle As EPenStyle = EPenStyle.psSolid, Optional WidthX As Long = 1, Optional WidthY As Long = 1, Optional Color As Long = vbBlack)
    With m_LogObj: .lopnStyle = PenStyle: .lopnWidthX = WidthX: .lopnWidthY = WidthY: .lopnColor = Color: End With
End Sub

Friend Sub NewPtr(ByVal lpLogObj As LongPtr)
    RtlMoveMemory ByVal Me.Ptr, ByVal lpLogObj, LenB(m_LogObj)
End Sub

'Copy to MNew
'Public Function Pen(Optional ByVal PenStyle As EPenStyle = EPenStyle.psSolid, Optional ByVal WidthX As Long = 1, Optional ByVal WidthY As Long = 1, Optional ByVal Color As Long = vbBlack) As Pen
'    Set Pen = New Pen: Pen.New_ PenStyle, WidthX, WidthY, Color
'End Function

Public Function NewC(Other As Pen)
    RtlMoveMemory m_LogObj, Other.Ptr, LenB(m_LogObj)
End Function

Public Sub Dispose():    If m_Handle Then DeleteObject m_Handle: m_Handle = 0
End Sub

Public Function Clone(Optional ByVal deep As Boolean = False) As Pen: Set Clone = New Pen: Clone.NewC Me: End Function

Public Sub SetParams(Optional PenStyle, Optional WidthX, Optional WidthY, Optional Color)
    If Not IsMissing(PenStyle) Then Me.PenStyle = CLng(PenStyle)
    If Not IsMissing(WidthX) Then Me.WidthX = CLng(WidthX)
    If Not IsMissing(WidthY) Then Me.WidthY = CLng(WidthY)
    If Not IsMissing(Color) Then Me.Color = CLng(Color)
End Sub

Public Function Create() As LongPtr: m_Handle = CreatePenIndirect(m_LogObj): Create = m_Handle: End Function

Public Property Get Ptr() As LongPtr:                   Ptr = VarPtr(m_LogObj):        End Property

Public Property Get Handle() As LongPtr:                Handle = m_Handle:             End Property

Public Property Get PenStyle() As EPenStyle:            PenStyle = m_LogObj.lopnStyle: End Property
Friend Property Let PenStyle(ByVal Value As EPenStyle): m_LogObj.lopnStyle = Value:    End Property

Public Property Get WidthX() As Long:                   WidthX = m_LogObj.lopnWidthX:  End Property
Friend Property Let WidthX(ByVal Value As Long):        m_LogObj.lopnWidthX = Value:   End Property

Public Property Get WidthY() As Long:                   WidthY = m_LogObj.lopnWidthY:  End Property
Friend Property Let WidthY(ByVal Value As Long):        m_LogObj.lopnWidthY = Value:   End Property

Public Property Get Color() As Long:                    Color = m_LogObj.lopnColor:    End Property
Friend Property Let Color(ByVal Value As Long):         m_LogObj.lopnColor = Value:    End Property


' v ' ############################## ' v '    all Enums    ' v ' ############################## ' v '
' v ' ############################## ' v '    EPenStyle    ' v ' ############################## ' v '
Public Function EPenStyle_ToStr(ByVal e As EPenStyle) As String
    Dim s As String
    Select Case e
    Case EPenStyle.psSolid:       s = "Solid      " & vbTab & "______"
    Case EPenStyle.psDash:        s = "Dash       " & vbTab & "------"
    Case EPenStyle.psDot:         s = "Dot        " & vbTab & "......"
    Case EPenStyle.psDashDot:     s = "DashDot    " & vbTab & "-.-.-."
    Case EPenStyle.psDashDotDot:  s = "DashDotDot " & vbTab & "-..-.."
    Case EPenStyle.psNull:        s = "Null       " & vbTab & "      "
    Case EPenStyle.psInsideFrame: s = "InsideFrame" & vbTab & "_"
    End With
    EPenStyle_ToStr = s
End Function

Public Function EPenStyle_Parse(ByVal s As String) As EPenStyle
    s = Trim(Split(s, vbTab)(0))
    Dim e As EPenStyle
    Select Case s
    Case "Solid":       e = EPenStyle.psSolid
    Case "Dash":        e = EPenStyle.psDash
    Case "Dot":         e = EPenStyle.psDot
    Case "DashDot":     e = EPenStyle.psDashDot
    Case "DashDotDot":  e = EPenStyle.psDashDotDot
    Case "Null":        e = EPenStyle.psNull
    Case "InsideFrame": e = EPenStyle.psInsideFrame
    End Select
    EPenStyle_Parse = e
End Function

Public Function EPenStyle_ToCBLB(ByVal CBLB As ComboBox) As EPenStyle
    CBLB.Clear: Dim i As Long: For i = 0 To 7: .AddItem EPenStyle_ToStr(i): Next
End Function




'class <GDI32.LOG...>:
'Private Sub Class_Initialize(): Me.New_: End Sub
'Private Sub Class_Terminate(): Me.Dispose: End Sub
'Friend Sub New_()
'Public Sub NewC(Other As <GDI32.LOG...>, optional byval deep as Boolean = False): End Sub
'Public Sub Dispose(): If m_Handle Then DeleteObject m_Handle: m_Handle = 0: End Sub
'Public Function Clone() As <GDI32.LOG...>: Set Clone = new <GDI32.LOG...>: Clone.NewC Me: End Function
'Public Sub SetParams(Optional Param1, Optional Param2, ...): End Sub
'Public Sub Create(): m_Handle = Create<GDI32.LOG...>: Create = m_Handle: End Sub
'Public Property Get Ptr() As LongPtr: Ptr = VarPtr(m_<GDI32.LOG...>): End Property
'Public Property Get Handle() As LongPtr: Handle = m_Handle: End Property




