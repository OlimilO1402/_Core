VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#Const IsUnicode = 1

'Thanks to Randy Birch
'Window Classes (CommCtrl.h)
'Private Const ANIMATE_CLASS      As String = "SysAnimate32"       ' Creates animation controls. These controls silently display an audio video interleaved (AVI) clip.
'Private Const DATETIMEPICK_CLASS As String = "SysDateTimePick32"  ' Creates date and time picker controls. These controls provide a simple interface to exchange date and time information with a user.
'Private Const HOTKEY_CLASS       As String = "msctls_hotkey32"    ' Creates hot key controls. These controls make it easy for the user to define hot keys.
'Private Const LINK_CLASS         As String = "SysLink"            ' Creates syslink controls. These controls are used with links.
'Private Const MONTHCAL_CLASS     As String = "SysMonthCal32"      ' Creates month calendar controls. These controls provide a simple way for a user to select a date from a familiar interface.
'Private Const NATIVEFNTCTL_CLASS As String = "???"                ' Creates native font controls. These controls are used with native fonts.
'Private Const PROGRESS_CLASS     As String = "msctls_progress32"  ' Creates progress bars. These controls indicate the progress of a lengthy operation.
'Private Const REBARCLASSNAME     As String = "ReBarWindow32"      ' Creates rebar controls. These controls act as a container for child windows.
'Private Const STANDARD_CLASSES   As String = "???"                ' Creates intrinsic User32 controls, such as button, edit, and scroll bar.
Private Const STATUSCLASSNAME    As String = "msctls_statusbar32" ' Creates status windows. These controls display status information in a horizontal window.
'Private Const TOOLBARCLASSNAME   As String = "ToolbarWindow"      ' Creates toolbars. These controls contain buttons that carry out menu commands.
'Private Const TOOLTIPS_CLASS     As String = "tooltips_class"     ' Creates tooltip controls. These controls display a small pop-up window containing a line of text that describes the purpose of a tool in an application.
'Private Const TRACKBAR_CLASS     As String = "msctls_trackbar32"  ' Creates trackbars. These controls let the user select from a range of values by moving a slider.
'Private Const UPDOWN_CLASS       As String = "msctls_updown32"    ' Creates up-down controls. These controls combine a pair of arrows with an edit control. Clicking the arrows increments or decrements the value in the edit control.
'Private Const WC_BUTTON          As String = "Button"             ' Creates button controls. These controls enable a user to initiate an action.
'Private Const WC_COMBOBOX        As String = "ComboBox"           ' Creates combobox   controls. These controls contain a drop-down list from which a single item can be selected.
'Private Const WC_COMBOBOXEX      As String = "ComboBoxEx32"       ' Creates ComboBoxEx controls. These controls provide an extension of the combo box control that provides native support for item images.
'Private Const WC_EDIT            As String = "Edit"               ' Creates edit controls. These controls contain editable text.
'Private Const WC_HEADER          As String = "SysHeader"          ' Creates header controls. These controls display headings at the top of columns of information and let the user sort the information by clicking the headings.
'Private Const WC_IPADDRESS       As String = "SysIPAddress32"     ' Creates IP address controls. These controls are similar to an edit control, but they enable you to enter a numeric address in Internet protocol (IP) format.
'Private Const WC_LINKA           As String = "SysLink"            ' Creates SysLink controls. These controls contain hypertext links.
'Private Const WC_LISTBOX         As String = "ListBox"            ' Creates listbox controls. These controls contain lists of items.
'Private Const WC_LISTVIEW        As String = "SysListView"        ' Creates list-view controls. These controls display a collection of items, each consisting of an icon and a label, and provide several ways to arrange the items.
'Private Const WC_NATIVEFONTCTL   As String = "NativeFontCtl"      ' Creates native font controls. These are invisible controls that work in the background to enable a dialog box's predefined controls to display the current system language.
'Private Const WC_PAGESCROLLER    As String = "SysPager"           ' Creates pager controls. These controls are used to contain and scroll another window.
'Private Const WC_SCROLLBAR       As String = "ScrollBar"          ' Creates scroll bar controls. These controls enable the user to scroll the contents of a window.
'Private Const WC_STATIC          As String = "Static"             ' Creates static controls. These controls contain noneditable text.
'Private Const WC_TABCONTROL      As String = "SysTabControl32"    ' Creates tab controls. These controls define multiple pages for the same area of a window or dialog box. Each page consists of a set of information or a group of controls that an application displays when the user selects the corresponding tab.
'Private Const WC_TREEVIEW        As String = "SysTreeView32"      ' Creates tree-view controls. These controls display a hierarchical list of items. Each item consists of a label and an optional bitmap.

Public Enum ESBPartAlignment ' Konstanten zur Ausrichtung von Grundflächen
    sbrLeft = 0        ' Links
    sbrCenter = 1      ' Zentriert
    sbrRight = 2       ' Rechts
End Enum

Public Enum ESBPartAutoSize  ' Konstanten zur automatischen Größenanpassung von Grundflächen
    sbrNoAutoSize = 0  ' Keine
    sbrSpring = 1      ' Gleitend
    sbrContents = 2    ' Inhalt
End Enum

Public Enum ESBPartBevel ' Konstanten für Grundflächenkanten
'    sbrNoBevel = 0 ' Keine
'    sbrInset = 1   ' Abgesenkt
'    sbrRaised = 2  ' Abgehoben
    sbrInset = 0       ' Abgesenkt, 'SBT_SUNKEN    =     0
    sbrNoBevel = 1 '&H1 '00 ' Keine      'SBT_NOBORDERS = &H100
    sbrRaised = 2 '&H2 '00  ' Abgehoben  'SBT_POPOUT    = &H200
End Enum

Public Enum ESBPartStyle ' Konstanten für Grundflächenstile
    sbrText = 0    ' Text
    sbrCaps = 1    ' FEST
    sbrNum = 2     ' NUM
    sbrIns = 3     ' Einfg
    sbrScrl = 4    ' ROLL
    sbrTime = 5    ' Zeit
    sbrDate = 6    ' Datum
    sbrKana = 7    ' Kana
End Enum


'Status Bar Styles
Private Const SBARS_SIZEGRIP    As Long = &H100 'Include a sizing grip at the right end of the status window.
Private Const WS_VISIBLE        As Long = &H10000000 'Creates a window that is initially visible.
Private Const WS_CHILD          As Long = &H40000000 'Creates a child window.
'uFlags text drawing operation values:
'also used by SB_SETTEXT, SB_SETTEXT, SB_GETTEXTLENGTH below

'does only work in ANSI
Private Const SBT_SUNKEN        As Long = &H0     ' Default 'The text is drawn with a border to appear lower than the plane of the window.
Private Const SBT_NOBORDERS     As Long = &H100   ' The text is drawn without borders.
Private Const SBT_POPOUT        As Long = &H200   ' The text is drawn with a border to appear higher 'than the plane of the window.
Private Const SBT_RTLREADING    As Long = &H400   ' SB_SETTEXT, SB_SETTEXT, SB_GETTEXTLENGTH flags only: 'Displays text using right-to-left reading order on 'Hebrew or Arabic systems.
Private Const SBT_TOOLTIPS      As Long = &H800   ' Statusbar can show Tooltips for each part
'Private Const SBT_NOTABPARSING  As Long = &H800
Private Const SBT_OWNERDRAW     As Long = &H1000  ' The text is drawn by the parent window.


Private Const SBB_HORIZONTAL    As Long = 0     ' horz border width
Private Const SBB_VERTICAL      As Long = 1     ' vert border width
Private Const SBB_DIVIDER       As Long = 2     ' vert part divider width

'Common Control Messages
Private Const CCM_FIRST            As Long = &H2000
Private Const CCM_SETBKCOLOR       As Long = (CCM_FIRST + 1)
Private Const CCM_SETCOLORSCHEME   As Long = (CCM_FIRST + 2)
Private Const CCM_GETCOLORSCHEME   As Long = (CCM_FIRST + 3)
Private Const CCM_GETDROPTARGET    As Long = (CCM_FIRST + 4)
Private Const CCM_SETUNICODEFORMAT As Long = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT As Long = (CCM_FIRST + 6)
Private Const CCM_SETVERSION       As Long = (CCM_FIRST + 7)
Private Const CCM_GETVERSION       As Long = (CCM_FIRST + 8)
Private Const CCM_SETNOTIFYWINDOW  As Long = (CCM_FIRST + 9)
Private Const CCM_SETWINDOWTHEME   As Long = (CCM_FIRST + &HB)
Private Const CCM_LAST             As Long = (CCM_FIRST + &H200)

Private Const WM_USER              As Long = &H400
'Status Bar Control Messages
'The following are only ANSI so they are obsolete
Private Const SB_SETTEXT           As Long = (WM_USER + 1)  ' The SB_SETTEXT message sets the text in the specified part of a status window. only used for Bevel
Private Const SB_GETTEXT           As Long = (WM_USER + 2)  ' The SB_GETTEXT message retrieves the text from the specified part of a status window. only used for Bevel
Private Const SB_GETTEXTLENGTH     As Long = (WM_USER + 3)  ' The SB_GETTEXTLENGTH message retrieves the length, in characters, of the text from the specified part of a status window.
Private Const SB_SETPARTS          As Long = (WM_USER + 4)  ' Sets the number of parts in a status window and the coordinate of the right edge of each part.
Private Const SB_GETPARTS          As Long = (WM_USER + 6)  ' Retrieves a count of the parts in a status window. The message also retrieves the coordinate of the right edge of the specified number of parts.
Private Const SB_GETBORDERS        As Long = (WM_USER + 7)  ' Retrieves the current widths of the horizontal and vertical borders of a status window.
Private Const SB_SETMINHEIGHT      As Long = (WM_USER + 8)  ' Sets the minimum height of a status window's drawing area.
Private Const SB_SIMPLE            As Long = (WM_USER + 9)  ' Specifies whether a status window displays simple text or displays all window parts set by a previous SB_SETPARTS message.
Private Const SB_GETRECT           As Long = (WM_USER + 10) ' Retrieves the bounding rectangle of a part in a status window.
Private Const SB_SETTEXTW          As Long = (WM_USER + 11) ' The SB_SETTEXT message sets the text in the specified part of a status window.
Private Const SB_GETTEXTLENGTHW    As Long = (WM_USER + 12) ' The SB_GETTEXTLENGTH message retrieves the length, in characters, of the text from the specified part of a status window.
Private Const SB_GETTEXTW          As Long = (WM_USER + 13) ' The SB_GETTEXT message retrieves the text from the specified part of a status window.
Private Const SB_ISSIMPLE          As Long = (WM_USER + 14) ' Checks a status bar control to determine if it is in simple mode.
Private Const SB_SETICON           As Long = (WM_USER + 15) ' Sets the icon for a part in a status bar.
Private Const SB_SETTIPTEXTW       As Long = (WM_USER + 17) ' Sets the tooltip text for a part in a status bar. The status bar must have been created with the SBT_TOOLTIPS style to enable tooltips.
Private Const SB_GETTIPTEXTW       As Long = (WM_USER + 19) ' Retrieves the tooltip text for a part in a status bar. The status bar must be created with the SBT_TOOLTIPS style to enable tooltips.
Private Const SB_GETICON           As Long = (WM_USER + 20) ' Retrieves the icon for a part in a status bar.
Private Const SB_SETBKCOLOR        As Long = CCM_SETBKCOLOR ' Sets the background color in a status bar.
Private Const SB_GETUNICODEFORMAT  As Long = CCM_SETUNICODEFORMAT ' Retrieves the Unicode character format flag for the control.
Private Const SB_SETUNICODEFORMAT  As Long = CCM_SETUNICODEFORMAT 'Sets the Unicode character format flag for the control. This message allows you to change the character set used by the control at run time rather than having to re-create the control.



Private Const SB_CONST_ALPHA    As Long = &H1
Private Const SB_GRAD_RECT      As Long = &H10
Private Const SB_GRAD_TRI       As Long = &H20
Private Const SB_SIMPLEID       As Long = &HFF

Private Const SB_HORZ           As Long = 0 'Scrollbar-constants?
Private Const SB_VERT           As Long = 1
Private Const SB_CTL            As Long = 2
Private Const SB_BOTH           As Long = 3

Private Const SB_LEFT           As Long = 6 'Scrollbar-constants?
Private Const SB_TOP            As Long = 6
Private Const SB_RIGHT          As Long = 7
Private Const SB_BOTTOM         As Long = 7
Private Const SB_ENDSCROLL      As Long = 8

Private Const SB_LINELEFT       As Long = 0
Private Const SB_LINERIGHT      As Long = 1
Private Const SB_LINEUP         As Long = 0
Private Const SB_LINEDOWN       As Long = 1

Private Const SB_NONE           As Long = &H0

Private Const SB_PAGELEFT       As Long = 2
Private Const SB_PAGERIGHT      As Long = 3
Private Const SB_PAGEUP         As Long = 2
Private Const SB_PAGEDOWN       As Long = 3

Private Const SB_PIXEL_ALPHA    As Long = &H2
Private Const SB_PREMULT_ALPHA  As Long = &H4

'Public Const SB_SETBKCOLOR As Long = CCM_SETBKCOLOR
Private Const SB_THUMBPOSITION  As Long = 4
Private Const SB_THUMBTRACK     As Long = 5



Private Const sbSimpleIdx       As Long = 255
Private Const WM_CREATE         As Long = &H1
Private Const WM_DESTROY        As Long = &H2
Private Const WM_SIZE           As Long = &H5
Private Const WM_SETTEXT        As Long = &HC
Private Const WM_GETTEXT        As Long = &HD
Private Const WM_GETTEXTLENGTH  As Long = &HE
Private Const WM_PAINT          As Long = &HF
Private Const WM_SETFONT        As Long = &H30
Private Const WM_GETFONT        As Long = &H31
Private Const WM_NCHITTEST      As Long = &H84

Private Const CCS_TOP           As Long = &H1&
Private Const CCS_NOMOVEY       As Long = &H2&
Private Const CCS_BOTTOM        As Long = &H3&
Private Const CCS_NORESIZE      As Long = &H4&
Private Const CCS_NOPARENTALIGN As Long = &H8&
Private Const CCS_NOHILITE      As Long = &H10&
Private Const CCS_ADJUSTABLE    As Long = &H20&
Private Const CCS_NODIVIDER     As Long = &H40&

'Also used w/ SB_GETRECT below
Private Type WinApiRECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

'#If IsUnicode Then
'    Private Declare Function CreateStatusWindow Lib "comctl32" Alias "CreateStatusWindowW" (ByVal Style As Long, ByVal lpszText As LongPtr, ByVal hWndParent As LongPtr, ByVal wID As Long) As LongPtr
'#Else
'    Private Declare Function CreateStatusWindow Lib "comctl32" Alias "CreateStatusWindowA" (ByVal Style As Long, ByVal lpszText As LongPtr, ByVal hWndParent As LongPtr, ByVal wID As Long) As LongPtr
'#End If

Private Declare Function CreateStatusWindowW Lib "comctl32" (ByVal Style As Long, ByVal lpszText As LongPtr, ByVal hWndParent As LongPtr, ByVal wID As Long) As LongPtr
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare Sub DrawStatusTextW Lib "comctl32" (ByVal hDC As LongPtr, lprc As WinApiRECT, ByVal pszText As LongPtr, uFlags As Long)
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, wParam As Any, lParam As Any) As LongPtr

'The data type for lpRect was changed from "As RECT" to
'"As Any" to allow a null pointer to be passed (i.e. ByRef 0)
'BTW the RECT structure is declared in MStatusDefs
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As WinApiRECT) As Long

Private Type StatusBarPart
    Alignment   As ESBPartAlignment
    AutoSize    As ESBPartAutoSize
    Bevel       As ESBPartBevel
    Style       As ESBPartStyle
    Enabled     As Boolean
    Visible     As Boolean
    Index       As Byte
    Left        As Single
    Width       As Single
    MinWidth    As Single
    Picture     As IPictureDisp 'IPicture, Picture, StdPicture
    Key         As String
    Text        As String
    ToolTipText As String
    Tag         As Variant
End Type

Private WithEvents mOwner As PictureBox
Attribute mOwner.VB_VarHelpID = -1
Private m_Handle  As LongPtr
Private m_Style   As Long
'Private m_Parts() As StatusBarPart
Private m_PartWidth() As Single
Private m_PartIcons() As StdPicture '
'Private m_PartBevel() As ESBPartBevel
'Negativer Wert -> SBPanel nimmt den restlich verbleibenden Teil der StatusBar ein
'Positiver Wert zwischen 0 und 1 -> Prozentualer Anteil der Gesamtbreite des Forms
'Positiver Wert größer 1 -> Länge in Pixel
'


Private Sub Class_Initialize()
    ReDim m_PartWidth(0)
    ReDim m_PartIcons(0)
    m_PartWidth(0) = -1 '50 ' mOwner.ScaleWidth / 2
End Sub

Private Sub Class_Terminate()
   If IsWindow(m_Handle) Then
      DestroyWindow m_Handle
      m_Handle = 0
      'EnableCtrls False
   End If
End Sub

Friend Sub New_(Owner As PictureBox) ', Parts() As Long)
    Set mOwner = Owner
    mOwner.Align = AlignConstants.vbAlignBottom
    mOwner.BorderStyle = 0 'BorderStyleConstants.vbTransparent
    mOwner.ScaleMode = ScaleModeConstants.vbPixels
    mOwner.Appearance = 0 '2D
    'Borderstyleconstants are originally Penstyleconstants
    'Mitglied von VBRUN.BorderStyleConstants
    'Strich (BorderStyle für Figur und Linie) (Shape and Line).
    'Const vbTransparent = 0
    'Const vbBSSolid = 1
    'Const vbBSDash = 2
    'Const vbBSDot = 3
    'Const vbBSDashDot = 4
    'Const vbBSDashDotDot = 5
    'Const vbBSInsideSolid = 6
    
    'SBT_SUNKEN    ' Default 'The text is drawn with a border to appear lower than the plane of the window.
    'SBT_NOBORDERS ' The text is drawn without borders.
    'SBT_POPOUT    ' The text is drawn with a border to appear higher 'than the plane of the window.
    
    m_Style = WS_VISIBLE Or WS_CHILD Or SBARS_SIZEGRIP Or SBT_TOOLTIPS ' 'Or SBT_POPOUT 'Or SBT_NOBORDERS '
    Dim txt As String: txt = "StatusBar..."
    'm_Handle = CreateStatusWindowW(m_Style, StrPtr(txt), mOwner.hWnd, 0)
    
    
'hwndStatus = CreateWindowEx(
'        0,                       // no extended styles
'        STATUSCLASSNAME,         // name of status bar class
'        (PCTSTR) NULL,           // no text when first created
'        SBARS_SIZEGRIP |         // includes a sizing grip
'        WS_CHILD | WS_VISIBLE,   // creates a visible child window
'        0, 0, 0, 0,              // ignores size and position
'        hwndParent,              // handle to parent window
'        (HMENU) idStatus,       // child window identifier
'        hinst,                   // handle to application instance
'        NULL);                   // no window creation data
    
    m_Handle = CreateWindowExW(0, ByVal StrPtr(STATUSCLASSNAME), ByVal StrPtr(txt), m_Style, 0, 0, 0, 0, mOwner.hWnd, ByVal 0, App.hInstance, ByVal 0)
    
    'Dim minh As Long: minh = mOwner.ScaleHeight '/ 3
    'SendMessageW m_Handle, SB_SETMINHEIGHT, ByVal minh, ByVal 0
    'SendMessageW m_Handle, WM_SIZE, 0, ByVal 0
    mOwner_Resize
End Sub

Private Sub mOwner_Resize()
    Me.MinHeight = mOwner.ScaleHeight
    MoveWindow m_Handle, 0, 0, mOwner.ScaleWidth, mOwner.ScaleHeight, True
    Me.Resize
End Sub

Public Property Get IsSimple() As Boolean
    Dim r As Long
    Dim b As Long, b2 As Long
    'r = SendMessageW(m_Handle, SB_ISSIMPLE, b, b2)
    r = SendMessageW(m_Handle, SB_SIMPLE, b, b2)
    Debug.Print r & " " & b & " " & b2
    IsSimple = b2
End Property

Public Property Let MinHeight(ByVal Value As Long)
    SendMessageW m_Handle, SB_SETMINHEIGHT, ByVal Value, ByVal 0
    SendMessageW m_Handle, WM_SIZE, 0, ByVal 0
End Property

Public Property Get MinHeight() As Long
    MinHeight = mOwner.ScaleHeight
End Property

Public Property Get Part_Count() As Byte
    'Returns the current number of existing parts in the status bar.
    'The SB_GETPARTS message also retrieve individual part
    'information. See the comments in Status.bas for more info.
    Part_Count = SendMessageW(m_Handle, SB_GETPARTS, 0, ByVal 0)
End Property

Public Property Let Part_Count(ByVal Value As Byte) '1-255 max!
    
    'Sets the specified number of status bar parts.
    'Any existing part with a greater index than the number of parts
    'specified by bParts is destroyed,
    'i.e 8 existing parts (0-7), 6 is specified for bParts,
    'the last 2 parts (6 & 7) are destroyed.
    
    'Array is zero based, will error back to
    'cmdDoMsgs_Click() if 0 is passed.
    
    ReDim Preserve m_PartWidth(0 To Value - 1)
    ReDim Preserve m_PartIcons(0 To Value - 1)
    Resize
'    'Set all but the last part so they have an equal width.
'    ReDim tmpParts(0 To Value - 1) As Long
'    Dim i As Byte
'    Dim w0 As Long: w0 = mOwner.ScaleWidth \ Value
'    Dim w As Long, l As Long
'    For i = 0 To Value - 2
'        If m_Parts(i).Width <= 0 Then m_Parts(i).Width = w0
'        w = m_Parts(i).Width
'        If m_Parts(i).Left = 0 Then m_Parts(i).Left = l
'        l = l + w
'        tmpParts(i) = l
'    Next
'
'    'Last part uses remaining real estate & extends to right edge.
'    m_Parts(Value - 1).Left = -1
'    tmpParts(Value - 1) = -1
'
'    SendMessageW m_Handle, SB_SETPARTS, ByVal Value, tmpParts(0)
    
End Property

Public Sub Resize(Optional ByVal SetDefaultPartWidth As Boolean = False)
    Dim scw As Single: scw = mOwner.ScaleWidth
    Dim i As Long, u As Long: u = UBound(m_PartWidth)
    Dim W As Single, c As Long: c = u + 1
    Dim dpw As Single: dpw = scw / c
    ReDim tmpParts(0 To u) As Long
    Dim L As Long
    For i = 0 To u
        If i = u Then
            m_PartWidth(i) = -1
        Else
            If SetDefaultPartWidth Then
                m_PartWidth(i) = dpw
            End If
        End If
        W = m_PartWidth(i)
        If 0 < W And W < 1 Then
            W = scw * W
        End If
        If W < 0 Then
            tmpParts(i) = -1
        Else
            L = L + W
            tmpParts(i) = L
        End If
    Next
    SendMessageW m_Handle, SB_SETPARTS, ByVal c, tmpParts(0)
End Sub

Public Property Get Part_TextLen(ByVal Index As Byte) As Long
    Dim rv As Long: rv = SendMessageW(m_Handle, SB_GETTEXTLENGTH, ByVal Index, ByVal 0)
    Part_TextLen = (rv And &HFFFF&)
End Property

Public Property Let Part_Bevel(ByVal Index As Byte, ByVal Value As ESBPartBevel)
'    'The Bevel property for the specified part is contained in the high word of the return value you get when asking for TEXTLENGTHW
    BevelText Index, Value, Me.Part_Text(Index)
'
'    Dim txt  As String: txt = Me.Part_Text(Index)
'    Dim wParam As Long: wParam = Index: LOWORD_HIBYTE(wParam) = Value
'    Dim rv As Long: rv = SendMessageW(m_Handle, SB_SETTEXTW, ByVal wParam, ByVal StrPtr(txt))
'                    rv = SendMessageW(m_Handle, WM_PAINT, ByVal 0, ByVal 0)
End Property

Public Property Get Part_Bevel(ByVal Index As Byte) As ESBPartBevel
    'The Bevel property for the specified part is contained in the high word of the return value you get when asking for TEXTLENGTHW
    Dim rv As Long: rv = SendMessageW(m_Handle, SB_GETTEXTLENGTHW, ByVal Index, ByVal 0)
    Part_Bevel = HIWORD_HIBYTE(rv) ' (rv And &HFFFF0000) \ &HFFFF&
End Property

Public Property Let Part_hIcon(ByVal Index As Byte, ByVal hIcon As LongPtr) ' StdPicture) 'or IPictureDisp ?
    Dim hr As LongPtr: hr = SendMessageW(m_Handle, SB_SETICON, ByVal Index, ByVal hIcon)
End Property

Public Property Get Part_hIcon(ByVal Index As Byte) As LongPtr 'StdPicture 'or IPictureDisp ?
    Dim hi As LongPtr: hi = SendMessageW(m_Handle, SB_GETICON, ByVal Index, ByVal 0)
    Part_hIcon = hi
End Property

Public Property Set Part_Icon(ByVal Index As Byte, ByVal Icon As StdPicture) ' IPictureDisp
    Set m_PartIcons(Index) = Icon: Part_hIcon(Index) = Icon.Handle
End Property

Public Property Get Part_Icon(ByVal Index As Byte) As StdPicture 'or IPictureDisp ?
    Set Part_Icon = m_PartIcons(Index)
End Property


'Static/Shared
Public Function ESBPartBevel_ToStr(ByVal e As ESBPartBevel) As String
    Dim s As String
    Select Case e
    Case ESBPartBevel.sbrInset:   s = "Inset"
    Case ESBPartBevel.sbrNoBevel: s = "NoBevel"
    Case ESBPartBevel.sbrRaised:  s = "Raised"
    End Select
    ESBPartBevel_ToStr = s
End Function

Public Property Get Part_Text(ByVal Index As Byte) As String
    Dim L As Long: L = Part_TextLen(Index)
    Dim Value As String: Value = Space(L) & vbNullChar
    L = SendMessageW(m_Handle, SB_GETTEXTW, ByVal Index, ByVal StrPtr(Value))
    Part_Text = Left(Value, L)
End Property

Public Property Let Part_Text(ByVal Index As Byte, ByVal ValueMax255Chars As String)
    
    'Sets the specified part's text.
    '
    'bPart:zero-based part to set, 255 = simple mode text.
    'wNewDrawOp:text drawing operation.
    'sText:text to set
    
    
    'Get the part's current drawing operation
    'before it might be updated below.
    'Dim wCurDrawOp As Integer: wCurDrawOp = GetCurDrawOp(Index, False)
    
    'Dim wNewDrawOp As Integer
    'Set the text w/ the drawing operation
    'SendMessageW m_Handle, SB_SETTEXT, ByVal Index Or wNewDrawOp, ByVal StrPtr(Value)
    
    
    'SendMessageW m_Handle, SB_SETTEXTW, ByVal Index, ByVal StrPtr(ValueMax255Chars)
    BevelText Index, Me.Part_Bevel(Index), ValueMax255Chars
    
    'Redraw the status bar only if the part's drawing
    'operation changed (reduces flicker).
    'If wCurDrawOp <> wNewDrawOp Then InvalidateRect m_Handle, ByVal 0, True
    
End Property

Private Sub BevelText(ByVal Index As Byte, ByVal BevelStyle As ESBPartBevel, ByVal Text As String)
    Dim wParam As Long: wParam = Index: LOWORD_HIBYTE(wParam) = BevelStyle
    Dim rv As Long: rv = SendMessageW(m_Handle, SB_SETTEXTW, ByVal wParam, ByVal StrPtr(Text))
                    rv = SendMessageW(m_Handle, WM_PAINT, ByVal 0, ByVal 0)
End Sub
'The tooltip text is displayed in two situations:
'* When the corresponding part in the status bar contains only an icon.
'* When the corresponding part in the status bar contains text that is truncated due to the size of the part.
Public Property Let Part_ToolTipText(ByVal Index As Byte, ByVal ValueMax255Chars As String)
    ' Sets the tooltip text for a part in a status bar. The status bar must have been created with the SBT_TOOLTIPS style to enable tooltips.
    Dim txt As String: txt = ValueMax255Chars & vbNullChar & vbNullChar
    SendMessageW m_Handle, SB_SETTIPTEXTW, ByVal Index, ByVal StrPtr(txt)
End Property

Public Property Get Part_ToolTipText(ByVal Index As Byte) As String
    ' Retrieves the tooltip text for a part in a status bar. The status bar must be created with the SBT_TOOLTIPS style to enable tooltips.
    Dim bufflen As Long:  bufflen = 255
    Dim buffer  As String: buffer = Space(bufflen) & vbNullChar & vbNullChar
    Dim wParam  As Long: wParam = Index: MPtr.HIWORD(wParam) = bufflen
    Dim hr As Long: hr = SendMessageW(m_Handle, SB_GETTIPTEXTW, ByVal wParam, ByVal StrPtr(buffer))
    Part_ToolTipText = Trim(buffer) ' Left(buffer, rv)
End Property

Public Property Get Part_Width(ByVal Index As Byte) As Single
    Part_Width = m_PartWidth(Index)
End Property

Public Property Let Part_Width(ByVal Index As Byte, ByVal Value As Single)
    m_PartWidth(Index) = Value
    Resize
End Property

Public Property Let PartsWidth(Values() As Single)
    m_PartWidth = Values
    Resize
End Property

Public Property Get PartsWidth() As Single()
    PartsWidth = m_PartWidth
End Property




'Private Function GetCurDrawOp(ByVal Index As Byte, Optional ByVal fRtnString As Boolean = False) ' As Integer 'String 'wieso String????
'
'   'Returns the current text drawing operation for the specified part.
'   '
'   'SB_GETTEXTLENGTH is used to determine the part's current
'   'drawing operation. SB_GETTEXT will rtn the exact same value,
'   'but requires a text buffer.
'   '
'   'When not in simple mode, SB_GETTEXTLENGTH  retrieves the
'   'text length for the part specified by bPart (0-254, 255 parts max).
'   'If in simple mode, SB_GETTEXTLENGTH will retrieve the simple
'   'mode text length *only* if bPart specifies any *valid* part index.
'   'The simple mode text length is NOT retrieved when bPart = 255
'   '(as is used to set text w/ SB_SETTEXT). Also applies to
'   'SB_GETTEXT.
'   '
'   'If fRtnString = True, returns the text drawing operation constant
'   'string. If False, returns the text drawing operation constant value.
'
'   Dim dwRtn As Long
'
'   dwRtn = SendMessageW(m_Handle, SB_GETTEXTLENGTHW, ByVal Index, 0)
'
'   'The text drawing operation for the specified
'   'part is contained in the high word of dwRtn.
'   dwRtn = (dwRtn And &HFFFF0000) \ &HFFFF&
'
'   If fRtnString Then
'
'     'Returning the string
'      Select Case dwRtn
'        Case SBT_SUNKEN:    GetCurDrawOp = "SBT_SUNKEN"
'        Case SBT_NOBORDERS: GetCurDrawOp = "SBT_NOBORDERS"
'        Case SBT_POPOUT:    GetCurDrawOp = "SBT_POPOUT"
'      End Select
'
'   Else
'
'     'Returning the value
'      GetCurDrawOp = dwRtn
'
'   End If
'End Function


'Class StatusBar
''Alle Mitglieder von ComctlLib.StatusBar
''Events:
'Event Click()                                               ' Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt.
'Event DblClick()                                            ' Tritt auf, wenn man eine Maustaste drückt und wieder losläßt und nochmals über einem Objekt drückt und wieder losläßt.
'Event DragDrop(Source As Control, x As Single, y As Single) ' Tritt auf, wenn eine Drag/Drop-Operation abgeschlossen ist.
'Event DragOver(Source As Control, x As Single, y As Single, State As Integer)  ' Tritt auf, während eine Drag/Drop-Operation durchgeführt wird.
'Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) ' Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat.
'Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) ' Tritt auf, wenn der Benutzer die Maus bewegt.
'Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)   ' Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat.
'Event OLECompleteDrag(Effect As Long)                                          ' OLECompleteDrag-Ereignis
'Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)                   ' OLEDragDrop-Ereignis
'Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) ' OLEDragOver-Ereignis
'Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)                                                                       ' OLEGiveFeedback-Ereignis
'Event OLESetData(Data As DataObject, DataFormat As Integer)                                                                            ' OLESetData-Ereignis
'Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)                                                                         ' OLEStartDrag-Ereignis
'Event PanelClick(Panel As Panel)    ' Ist dem Standardereignis Click ähnlich, tritt aber auf, wenn ein Benutzer eine Maustaste
'                                    ' über einem der Panel-Objekte des Statusleiste-Steuerelements drückt und wieder losläßt.
'Event PanelDblClick(Panel As Panel) ' Ist dem Standardereignis Click ähnlich, tritt aber auf, wenn ein Benutzer eine Maustaste
'                                    ' über einem der Panel-Objekte des Statusleiste-Steuerelements zweimal drückt und wieder losläßt.
''Subs
'Sub Drag([Action])      ' Dient zum Einleiten, Beenden oder Abbrechen einer Ziehoperation auf einem beliebigen Objekt außer den Steuerelementen Linie, Menü, Figur und Zeitgeber.
'Sub Move(Left As Single, [Top], [Width], [Height]) ' Verschiebt ein Objekt.
'Sub OLEDrag()           ' Startet ein OLE-Drag/Drop-Ereignis mit dem angegebenen Steuerelement als Quelle.
'Sub Refresh()           ' Erzwingt ein vollständiges Neuzeichnen eines Formulars oder Steuerelements.
'Sub ShowWhatsThis()     ' Zeigt ein ausgewähltes Thema in einer Hilfedatei über die von der Windows 95-Hilfe bereitgestellte Direkthilfe an.
'Sub ZOrder([Position])  ' Legt ein angegebenes Objekt an der ersten oder letzten Stelle der Z-Reihenfolge innerhalb der zugehörigen Grafikstufe ab.
'
''Properties
'Property Align           As Integer                 ' Gibt einen Wert zurück oder legt einen Wert fest, der bestimmt, ob ein Objekt in einem Formular angezeigt wird.
'Property Container       As Object                  ' Gibt den Container eines Objekts zurück.
'Property DragIcon        As IPictureDisp            ' Gibt das Symbol zurück, das als Zeiger in einer Drag/Drop-Operation angezeigt werden soll, oder legt dieses fest.
'Property DragMode        As Integer                 ' Gibt einen Wert zurück, der bestimmt, ob der manuelle oder der automatische Ziehmodus verwendet wird, oder legt diesen fest.
'Property Enabled         As Boolean                 ' Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest.
'Property Font            As IFontDisp               ' Gibt ein Font-Objekt zurück.
'Property Height          As Single                  ' Gibt die Höhe eines Objekts zurück oder legt diese fest.
'Property hWnd            As OLE_HANDLE              ' Gibt eine Zugriffsnummer auf ein Formular oder Steuerelement zurück.
'Property Index           As Integer                 ' schreibgeschützt    Gibt die Zahl zurück, die ein Steuerelement in einem Steuerelementefeld identifiziert, oder legt diese fest.
'Property Left            As Single                  ' Gibt die Entfernung zwischen dem internen linken Rand eines Objekts und dem linken Rand des zugehörigen Containers zurück oder legt diese fest.
'Property MouseIcon       As IPictureDisp            ' Legt ein benutzerdefiniertes Maussymbol fest.
'Property MousePointer    As MousePointerConstants   ' Gibt zurück oder legt fest, welche Art von Mauszeiger angezeigt wird, wenn er sich über einem Teil eines Objekts befindet.
'Property Name            As String                  ' schreibgeschützt'    Gibt den Namen zurück, der im Code zur Identifikation eines Objekts verwendet wird.
'Property Object          As Object                  ' schreibgeschützt'    Gibt ein Objekt in einem Steuerelement zurück.
'Property OLEDropMode     As OLEDropConstants        ' Gibt zurück oder legt fest, ob dieses Steuerelement ein Ziel für das Ablegen von OLE-Objekten darstellen kann.
'Property Panels          As Panels                  ' Gibt einen Verweis auf eine Auflistung von Panel-Objekten zurück.
'Property Parent          As Object                  ' schreibgeschützt'    Gibt das Objekt zurück, auf dem sich dieses Objekt befindet.
'Property ShowTips        As Boolean                 ' Aktiviert/Deaktiviert die QuickInfo für Grundflächen
'Property SimpleText      As String                  ' Gibt zurück oder legt fest, welcher Text angezeigt wird, wenn die Style-Eigenschaft eines Statusleiste-Steuerelements auf 'Einfach' gesetzt wird.
'Property Style           As SbarStyleConstants      ' Gibt den Stil für einzelne (einfache) oder mehrfache Grundflächen zurück oder legt ihn fest
'Property TabIndex        As Integer                 ' Gibt die Aktivierreihenfolge eines Objekts innerhalb des zugehörigen übergeordneten Formulars zurück oder legt diese fest.
'Property Tag             As String                  ' Speichert alle zusätzlichen Daten, die Sie für Ihr Programm benötigen.
'Property ToolTipText     As String                  ' Gibt den Text zurück, der angezeigt wird, wenn die Maus über dem Steuerelement verweilt, oder legt den Text fest.
'Property Top             As Single                  ' Gibt die Entfernung zwischen dem internen oberen Rand eines Objekts und dem oberen Rand des zugehörigen Containers zurück oder legt diesen fest.
'Property Visible         As Boolean                 ' Gibt einen Wert zurück, der bestimmt, ob ein Objekt sichtbar oder versteckt ist, oder legt diesen fest.
'Property WhatsThisHelpID As Long                    ' Gibt eine zugeordnete Kontextnummer für ein Objekt zurück oder legt diese fest.
'Property Width           As Single                  ' Gibt die Breite eines Objekts zurück oder legt diese fest.
'
'Class Panel ' Mitglied von ComctlLib ' Statusleiste-Grundfläche
' Alle Mitglieder von ComctlLib.Panel
'Property Alignment       As PanelAlignmentConstants ' Gibt die Ausrichtung eines Statusleiste-Steuerelements (unten, oben, links, rechts oder unverankert) zurück oder legt sie fest
'Property AutoSize        As PanelAutoSizeConstants  ' Gibt einen Wert zurück oder legt einen Wert fest, der zuläßt, daß die Breite eines Panel-Objekts automatisch
'                                                    ' angepaßt werden kann, wenn sich der Grundflächeninhalt ändert oder die Größe des übergeordneten Formulars ändert
'Property Bevel           As PanelBevelConstants     ' Gibt die Kantendarstellung eines Panel-Objekts (abgehoben, abgesenkt oder keine Kante) zurück oder legt sie fest
'Property Enabled         As Boolean                 ' Legt einen Wert fest oder gibt einen Wert zurück, der bestimmt, ob ein Formular oder Steuerelement auf vom Benutzer generierte Ereignisse reagieren kann
'Property Index           As Integer                 ' Gibt den Index eines Objekts in einer Auflistung zurück oder legt ihn fest. Zur Laufzeit schreibgeschützt.
'Property Key             As String                  ' Legt die eindeutige Zeichenfolge eines Objekts in einer Auflistung fest oder gibt die Zeichenfolge zurück
'Property Left            As Single                  ' Gibt den Abstand zwischen dem internen linken Rand eines Objekts und dem linken Rand des zugehörigen Containers zurück oder legt ihn fest
'Property MinWidth        As Single                  ' Gibt die minimale Breite eines Panel-Objekts zurück oder legt sie fest
'Property Picture         As IPictureDisp            ' Gibt die Picture-Eigenschaft für ein Panel-Objekt zurück oder legt sie fest
'Property Style           As PanelStyleConstants     ' Gibt den Stil eines Panel-Objekts eines Statusleiste-Steuerelements zurück oder legt ihn fest
'Property Tag             As Variant                 ' Speichert alle zusätzlichen Daten, die Sie für Ihr Programm benötigen.
'Property Text            As String                  ' Gibt den in einem Steuerelement anzuzeigenden Text zurück oder legt ihn fest
'Property ToolTipText     As String                  ' Legt die Zeichenfolge fest, die angezeigt wird, wenn ein Cursor über einem Objekt bewegt wird, oder gibt diese Zeichenfolge zurück.
'Property Visible         As Boolean                 ' Gibt einen Wert zurück oder legt einen Wert fest, der bestimmt, ob ein Objekt sichtbar oder ausgeblendet ist
'Property Width           As Single                  ' Gibt die Breite eines Objekts zurück oder legt sie fest
'
'Enum PanelAlignmentConstants ' Konstanten zur Ausrichtung von Grundflächen
'    Const sbrLeft = 0        ' Links
'    Const sbrCenter = 1      ' Zentriert
'    Const sbrRight = 2       ' Rechts
'End Enum
'
'Enum PanelAutoSizeConstants  ' Konstanten zur automatischen Größenanpassung von Grundflächen
'    Const sbrNoAutoSize = 0  ' Keine
'    Const sbrSpring = 1      ' Gleitend
'    Const sbrContents = 2    ' Inhalt
'End Enum
'
'Enum PanelBevelConstants ' Konstanten für Grundflächenkanten
'    Const sbrNoBevel = 0 ' Keine
'    Const sbrInset = 1   ' Abgesenkt
'    Const sbrRaised = 2  ' Abgehoben
'End Enum
'
'Enum PanelStyleConstants ' Konstanten für Grundflächenstile
'    Const sbrText = 0    ' Text
'    Const sbrCaps = 1    ' FEST
'    Const sbrNum = 2     ' NUM
'    Const sbrIns = 3     ' Einfg
'    Const sbrScrl = 4    ' ROLL
'    Const sbrTime = 5    ' Zeit
'    Const sbrDate = 6    ' Datum
'    Const sbrKana = 7    ' Kana
'End Enum
'
'Public Enum ESBPartAlignment ' Konstanten zur Ausrichtung von Grundflächen
'    Const sbrLeft = 0        ' Links
'    Const sbrCenter = 1      ' Zentriert
'    Const sbrRight = 2       ' Rechts
'End Enum
'
'Public Enum ESBPartAutoSize  ' Konstanten zur automatischen Größenanpassung von Grundflächen
'    Const sbrNoAutoSize = 0  ' Keine
'    Const sbrSpring = 1      ' Gleitend
'    Const sbrContents = 2    ' Inhalt
'End Enum
'
'Public Enum ESBPartBevel ' Konstanten für Grundflächenkanten
'    Const sbrNoBevel = 0 ' Keine
'    Const sbrInset = 1   ' Abgesenkt
'    Const sbrRaised = 2  ' Abgehoben
'End Enum
'
'Public Enum ESBPartStyle ' Konstanten für Grundflächenstile
'    Const sbrText = 0    ' Text
'    Const sbrCaps = 1    ' FEST
'    Const sbrNum = 2     ' NUM
'    Const sbrIns = 3     ' Einfg
'    Const sbrScrl = 4    ' ROLL
'    Const sbrTime = 5    ' Zeit
'    Const sbrDate = 6    ' Datum
'    Const sbrKana = 7    ' Kana
'End Enum

'Private Function StatusBarPart(ByVal Alignment As ESBPartAlignment, ByVal AutoSize As ESBPartAutoSize, ByVal Bevel As ESBPartBevel, ByVal Style As ESBPartStyle, _
'                               ByVal Enabled As Boolean, ByVal Visible As Boolean, ByVal Index As Byte, ByVal Left As Single, ByVal Width As Single, MinWidth As Single, _
'                               Picture As StdPicture, ByVal Key As String, ByVal Text As String, ByVal ToolTipText As String, Tag As Variant) As StatusBarPart
'    With StatusBarPart
'        .Alignment = Alignment: .AutoSize = AutoSize: .Bevel = Bevel: .Style = Style: .Enabled = Enabled: .Visible = Visible: .Index = Index
'        .Left = Left: .Width = Width: .MinWidth = MinWidth: Set .Picture = Picture: .Key = Key: .Text = Text: .ToolTipText = ToolTipText: .Tag = Tag
'    End With
'End Function

'' v ' ############################## ' v '     all mOwner-Events (PictureBox)    ' v ' ############################## ' v '
'
'Private Sub mOwner_Click()
'    MsgBox "User Clicked mOwner!"
'End Sub
'
'Private Sub mOwner_DblClick()
'    '
'End Sub
'
'Private Sub mOwner_GotFocus()
'    '
'End Sub
'
'Private Sub mOwner_KeyDown(KeyCode As Integer, Shift As Integer)
'    '
'End Sub
'
'Private Sub mOwner_KeyPress(KeyAscii As Integer)
'    '
'End Sub
'
'Private Sub mOwner_KeyUp(KeyCode As Integer, Shift As Integer)
'    '
'End Sub
'
'Private Sub mOwner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '
'End Sub
'
'Private Sub mOwner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '
'End Sub
'
'Private Sub mOwner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '
'End Sub
'
'Private Sub mOwner_OLECompleteDrag(Effect As Long)
'    '
'End Sub
'
'Private Sub mOwner_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '
'End Sub
'
'Private Sub mOwner_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
'    '
'End Sub
'
'Private Sub mOwner_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
'    '
'End Sub
'
'Private Sub mOwner_OLESetData(Data As DataObject, DataFormat As Integer)
'    '
'End Sub
'
'Private Sub mOwner_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
'    '
'End Sub
'
