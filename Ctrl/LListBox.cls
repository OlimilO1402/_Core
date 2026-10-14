VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements ListBox

Private Const GWL_WNDPROC     As Long = (-4&)
Private Const SRCCopy         As Long = &HCC0020

Private Const COLOR_HIGHLIGHT As Long = &HD
Private Const COLOR_HIGHLIGHTTEXT As Long = &HE
Private Const COLOR_BTNTEXT   As Long = &H12

Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_RAISEDINNER As Long = &H4
Private Const EDGE_RAISED     As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)

Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_VSCROLL       As Long = &H200000
Private Const WS_BORDER        As Long = &H800000
Private Const WS_VISIBLE       As Long = &H10000000
Private Const WS_CHILD         As Long = &H40000000

Private Const LBS_NOTIFY            As Long = &H1&
Private Const LBS_SORT              As Long = &H2&
Private Const LBS_NOREDRAW          As Long = &H4&
Private Const LBS_MULTIPLESEL       As Long = &H8&

Private Const LBS_OWNERDRAWFIXED    As Long = &H10&
Private Const LBS_OWNERDRAWVARIABLE As Long = &H20&
Private Const LBS_HASSTRINGS        As Long = &H40&
Private Const LBS_USETABSTOPS       As Long = &H80&

Private Const LBS_NOINTEGRALHEIGHT  As Long = &H100&
Private Const LBS_MULTICOLUMN       As Long = &H200&
Private Const LBS_WANTKEYBOARDINPUT As Long = &H400&
Private Const LBS_EXTENDEDSEL       As Long = &H800&

Private Const LBS_DISABLENOSCROLL   As Long = &H1000&
Private Const LBS_NODATA            As Long = &H2000&
Private Const LBS_NOSEL             As Long = &H4000&
Private Const LBS_STANDARD          As Long = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)

'Messages
Private Const LB_CTLCODE             As Long = 0&
Private Const LB_ADDSTRING           As Long = &H180 ' Adds a string to a list box.
Private Const LB_INSERTSTRING        As Long = &H181 ' Inserts a string or item data into a list box. Unlike the LB_ADDSTRING message, the LB_INSERTSTRING message does not cause a list with the LBS_SORT style to be sorted.
Private Const LB_DELETESTRING        As Long = &H182 ' Deletes a string in a list box.
Private Const LB_SELITEMRANGEEX      As Long = &H183 ' Selects one or more consecutive items in a multiple-selection list box.
Private Const LB_RESETCONTENT        As Long = &H184 ' Removes all items from a list box.
Private Const LB_SETSEL              As Long = &H185 ' Selects a string in a multiple-selection list box.
Private Const LB_SETCURSEL           As Long = &H186 ' Selects a string and scrolls it into view, if necessary.
Private Const LB_GETSEL              As Long = &H187 ' Gets the selection state of an item.
Private Const LB_GETCURSEL           As Long = &H188 ' Gets the index of the currently selected item, if any, in a single-selection list box.
Private Const LB_GETTEXT             As Long = &H189 ' Gets a string from a list box.
Private Const LB_GETTEXTLEN          As Long = &H18A ' Gets the length of a string in a list box.
Private Const LB_GETCOUNT            As Long = &H18B ' Gets the number of items in a list box.
Private Const LB_SELECTSTRING        As Long = &H18C ' Searches a list box for an item that begins with the characters in a specified string.
Private Const LB_DIR                 As Long = &H18D ' Adds names to the list displayed by a list box.
Private Const LB_GETTOPINDEX         As Long = &H18E ' Gets the index of the first visible item in a list box.
Private Const LB_FINDSTRING          As Long = &H18F ' Finds the first string in a list box that begins with the specified string.
Private Const LB_GETSELCOUNT         As Long = &H190 ' Gets the total number of selected items in a multiple-selection list box.
Private Const LB_GETSELITEMS         As Long = &H191 ' Fills a buffer with an array of integers that specify the item numbers of selected items in a multiple-selection list box.
Private Const LB_SETTABSTOPS         As Long = &H192 ' Sets the tab-stop positions in a list box.
Private Const LB_GETHORIZONTALEXTENT As Long = &H193 ' Gets the width, in pixels, that a list box can be scrolled horizontally (the scrollable width) if the list box has a horizontal scroll bar.
Private Const LB_SETHORIZONTALEXTENT As Long = &H194 ' Sets the width, in pixels, by which a list box can be scrolled horizontally (the scrollable width).
Private Const LB_SETCOLUMNWIDTH      As Long = &H195 ' Sets the width, in pixels, of all columns in a multiple-column list box.
Private Const LB_ADDFILE             As Long = &H196 ' Adds the specified file name to a list box that contains a directory listing.
Private Const LB_SETTOPINDEX         As Long = &H197 ' Ensures that the specified item in a list box is visible.
Private Const LB_GETITEMRECT         As Long = &H198 ' Gets the dimensions of the rectangle that bounds a list box item as it is currently displayed in the list box.
Private Const LB_GETITEMDATA         As Long = &H199 ' Gets the application-defined value associated with the specified list box item.
Private Const LB_SETITEMDATA         As Long = &H19A ' Sets a value that is associated with the specified item in a list box.
Private Const LB_SELITEMRANGE        As Long = &H19B ' Selects or deselects one or more consecutive items in a multiple-selection list box.
Private Const LB_SETANCHORINDEX      As Long = &H19C ' Sets the anchor item that is, the item from which a multiple selection starts. A multiple selection spans all items from the anchor item to the caret item.
Private Const LB_GETANCHORINDEX      As Long = &H19D ' Gets the index of the anchor item that is, the item from which a multiple selection starts.
Private Const LB_SETCARETINDEX       As Long = &H19E ' Sets the focus rectangle to the item at the specified index in a multiple-selection list box. If the item is not visible, it is scrolled into view.
Private Const LB_GETCARETINDEX       As Long = &H19F ' Retrieves the index of the item that has the focus rectangle in a multiple-selection list box. The item may or may not be selected.
Private Const LB_SETITEMHEIGHT       As Long = &H1A0 ' Sets the height, in pixels, of items in a list box.
Private Const LB_GETITEMHEIGHT       As Long = &H1A1 ' Gets the height of items in a list box.
Private Const LB_FINDSTRINGEXACT     As Long = &H1A2 ' Finds the first list box string that exactly matches the specified string, except that the search is not case sensitive.
'Private Const LB_                    As Long = &H1A3 '
'Private Const LB_                    As Long = &H1A4 '
Private Const LB_SETLOCALE           As Long = &H1A5 ' Sets the current locale of the list box.
Private Const LB_GETLOCALE           As Long = &H1A6 ' Gets the current locale of the list box.
Private Const LB_SETCOUNT            As Long = &H1A7 ' Sets the count of items in a list box created with the LBS_NODATA style and not created with the LBS_HASSTRINGS style.
Private Const LB_INITSTORAGE         As Long = &H1A8 ' Allocates memory for storing list box items. This message is used before an application adds a large number of items to a list box.
Private Const LB_ITEMFROMPOINT       As Long = &H1A9 ' Gets the zero-based index of the item nearest the specified point in a list box.

Private Const LB_MSGMAX_4            As Long = &H1B0
Private Const LB_MSGMAX_WCE4         As Long = &H1B1
Private Const LB_MULTIPLEADDSTRING   As Long = &H1B1 '
Private Const LB_GETLISTBOXINFO      As Long = &H1B2 ' Gets the number of items per column in a specified list box.
Private Const LB_MSGMAX_501          As Long = &H1B3

Private Const ODS_SELECTED     As Long = &H1

Private Const WM_SETREDRAW       As Long = &HB
Private Const WM_SETFONT         As Long = &H30
Private Const WM_GETFONT         As Long = &H31
Private Const WM_DRAWITEM        As Long = &H2B
Private Const WM_CTLCOLORLISTBOX As Long = &H134
Private Const WM_LBUTTONDOWN     As Long = &H201
Private Const WM_LBUTTONUP       As Long = &H202

Private Const DT_LEFT            As Long = &H0
Private Const DT_VCENTER         As Long = &H4
Private Const DT_SINGLELINE      As Long = &H20

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType    As Long
    CtlID      As Long
    itemID     As Long
    itemAction As Long
    itemState  As Long
    hwndItem   As Long
    hDC        As Long
    rcItem     As RECT
    ItemData   As Long
End Type

Private Type DRAWTEXTPARAMS
    cbSize        As Long
    iTabLength    As Long
    iLeftMargin   As Long
    iRightMargin  As Long
    uiLengthDrawn As Long
End Type

Private Type ODLBTYPE
    Forecolor         As Long
    Backcolor         As Long
    ForeColorSelected As Long
    BackColorSelected As Long
    BorderStyle       As Long
    DrawStyle         As Long
    ItemHeight        As Long
    hWnd              As Long
    Left              As Long
    Top               As Long
    Width             As Long
    Height            As Long
    hFont             As Long
    Parent            As PictureBox 'Form
End Type

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-sendmessage
'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-sendmessagea
'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-sendmessagew
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Rückgabewert: LRESULT; Der Rückgabewert gibt das Ergebnis der Nachrichtenverarbeitung an. dies hängt von der gesendeten Nachricht ab.

Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal ByteLength As Long)
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long



'UIPI: User Interface Privilege Isolation (UIPI) is a technology introduced in Windows Vista and Windows Server 2008 to combat shatter attack exploits.

Private WithEvents mListBox As ListBox
Attribute mListBox.VB_VarHelpID = -1
Private m_hWnd    As Long
Private m_HighlightEntries As Boolean

Public Event Click()
Public Event DblClick()
Public Event GotFocus()
Public Event LostFocus()
Public Event Scroll()

Private Sub Class_Initialize()
    '
End Sub

Friend Sub New_(aListBox As ListBox)
    Set mListBox = aListBox
    m_hWnd = mListBox.hWnd
End Sub

Private Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Sub Clear()
    mListBox.Clear
End Sub

Public Property Get Enabled() As Boolean
    Enabled = mListBox.Enabled
End Property
Public Property Let Enabled(ByVal Value As Boolean)
    mListBox.Enabled = Value
End Property

Public Property Get Visible() As Boolean
    Visible = mListBox.Visible
End Property
Public Property Let Visible(ByVal Value As Boolean)
    mListBox.Visible = Value
End Property

Public Property Get HighlightEntriesOnMouseMove() As Boolean
    HighlightEntriesOnMouseMove = m_HighlightEntries
End Property
Public Property Let HighlightEntriesOnMouseMove(ByVal Value As Boolean)
    m_HighlightEntries = Value
End Property

' Sucht String in ListBox, gibt Index zurück
' wenn Übereinstimmung gefunden, sonst -1
' Der Eintrag wird Markiert
Public Function SelectItem(ByVal FindString As String, ByVal StartIndex As Long) As Long
    SelectItem = SendMessage(hWnd, LB_SELECTSTRING, 0, ByVal FindString)
End Function

' Sucht exakten String in ListBox, gibt Index bei
' Übereinstimmung zurück, sonst -1
' Der Eintrag wird nicht markiert
Public Function FindExactItem(ByVal FindString As String, ByVal StartIndex As Long) As Long
    FindExactItem = SendMessage(hWnd, LB_FINDSTRINGEXACT, StartIndex, ByVal FindString)
End Function

' Sucht (nicht exakten) String in ListBox, gibt Index bei
' Übereinstimmung zurück, sonst -1
' Der Eintrag wird nicht Markiert
Public Function FindItem(FindString As String, StartIndex As Long) As Long
    FindItem = SendMessage(hWnd, LB_FINDSTRING, StartIndex, ByVal FindString)
End Function

' ListIndex von ListBox setzen
' setzt den ListIndex und markiert den Eintrag
' Ein Click wird nicht ausgelöst!
Public Property Get ListIndex() As Long
    ListIndex = SendMessage(hWnd, LB_GETCURSEL, 0, ByVal 0&)
End Property
' ListIndex von ListBox setzen
' setzt den ListIndex und markiert den Eintrag
' Ein Click wird nicht ausgelöst!
Public Property Let ListIndex(ByVal Index As Long)
    ListIndex = SendMessage(hWnd, LB_SETCURSEL, Index, ByVal 0&)
End Property

' Anzahl List-Einträge ermitteln (ListCount)
Public Property Get ListCount() As Long
    ListCount = SendMessage(hWnd, LB_GETCOUNT, 0, ByVal 0&)
End Property

' Liest den Text in einer ListBox aus mit dem
' angegebenen Index aus und gibt diesen zurück
Public Property Get Item(ByVal Index As Long) As String
    Dim TextLen As Long: TextLen = SendMessage(hWnd, LB_GETTEXTLEN, Index, ByVal 0&)
    Dim Buffer As String
    If TextLen > 0 Then
        Buffer = Space(TextLen)
        SendMessage hWnd, LB_GETTEXT, Index, ByVal Buffer
    End If
    Item = Buffer
End Property

' Ersetzt den Item-Text an der Position 'Index' mit 'NewText'
' Der Eintrag wird markiert!
Public Property Let Item(ByVal Index As Long, NewText As String)
    ' zunächst Eintrag löschen
    If SendMessage(hWnd, LB_DELETESTRING, Index, ByVal 0&) > 0 Then
        ' und neuen Eintrag an diese Position wieder einfügen
        SendMessage hWnd, LB_INSERTSTRING, Index, ByVal NewText
        ' und noch selektieren
        SendMessage hWnd, LB_SETCURSEL, Index, ByVal 0
    End If
End Property

' Ermittelt den benutzerdefinierten Wert ItemData
' Der Wert wird mit LBSetItemData() zugewiesen
Public Property Get ItemData(ByVal Index As Long) As Long
    ItemData = SendMessage(hWnd, LB_GETITEMDATA, Index, ByVal 0&)
End Property

' setzt den benutzerdefinierten Wert ItemData
Public Property Let ItemData(ByVal Index As Long, ByVal ItemData As Long)
    Call SendMessage(hWnd, LB_SETITEMDATA, Index, ByVal ItemData)
End Property

' Löscht den Text aus der ListBox
' Gibt bei Erfolg die Anzahl der verbleibenden
' Einträge in der Liste zurück, bei Fehler -1
Public Function Delete(ByVal Index As Long) As Long
    Delete = SendMessage(hWnd, LB_DELETESTRING, Index, ByVal 0&)
    SendMessage hWnd, LB_SETCURSEL, Index, ByVal 0
End Function
 
' Fügt ein Eintrag an die Position 'Index' ein
' Der eingefügte Text wird nicht matkiert.
Public Sub Insert(ByVal Index As Long, ByVal Text As String)
    SendMessage hWnd, LB_INSERTSTRING, Index, ByVal Text
    SendMessage hWnd, LB_SETCURSEL, Index, ByVal 0
End Sub

' Hängt ans Ende der Listbox einen Eintrag an.
' Liefert die Position zurück, an der der text eingefügt wurde
' Der hinzugefügte Text wird nicht markiert
Public Function AddItem(ByVal Text As String) As Long
    AddItem = SendMessage(hWnd, LB_ADDSTRING, 0, ByVal Text)
End Function

Public Sub AddList(Texts() As String)
    SendMessage hWnd, LB_MULTIPLEADDSTRING, 0, Texts(0)
End Sub

' ############################## '     ListBox Events    ' ############################## '
Private Sub mListBox_Click()
    RaiseEvent Click
End Sub

Private Sub mListBox_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub mListBox_DragDrop(Source As Control, X As Single, Y As Single)
    '
End Sub

Private Sub mListBox_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    '
End Sub

Private Sub mListBox_GotFocus()
    RaiseEvent GotFocus
End Sub

Private Sub mListBox_ItemCheck(Item As Integer)
    '
End Sub

Private Sub mListBox_KeyDown(KeyCode As Integer, Shift As Integer)
    '
End Sub

Private Sub mListBox_KeyPress(KeyAscii As Integer)
    '
End Sub

Private Sub mListBox_KeyUp(KeyCode As Integer, Shift As Integer)
    '
End Sub

Private Sub mListBox_LostFocus()
    RaiseEvent LostFocus
End Sub

Private Sub mListBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
End Sub

'VB 5/6-Tipp 0052: Highlighting für Listboxeinträge
'https://www.activevb.de/tipps/vb6tipps/tipp0052.html
Private Sub mListBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_HighlightEntries Then Exit Sub
    Dim SX As Single: SX = List1.Parent.ScaleX(X, List1.Parent.ScaleMode, vbPixels)
    Dim SY As Single: SY = List1.Parent.ScaleY(Y, List1.Parent.ScaleMode, vbPixels)
    Dim XY As Long:   XY = CLng(SX) + &H10000 * CLng(SY)
    Dim li As Long:   li = SendMessage(List1.hWnd, LB_ITEMFROMPOINT, 0, ByVal XY)
    If li < List1.ListCount Then List1.ListIndex = li
End Sub

Private Sub mListBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
End Sub

Private Sub mListBox_OLECompleteDrag(Effect As Long)
    '
End Sub

Private Sub mListBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
End Sub

Private Sub mListBox_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    '
End Sub

Private Sub mListBox_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    '
End Sub

Private Sub mListBox_OLESetData(Data As DataObject, DataFormat As Integer)
    '
End Sub

Private Sub mListBox_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    '
End Sub

Private Sub mListBox_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub mListBox_Validate(Cancel As Boolean)
    '
End Sub
