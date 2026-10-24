VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBcrud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mBtnAdd    As CommandButton ' 1. [ + ]  Hinzufügen,  ein neues       Objekt Erstellen, Editieren, und am Ende der Liste anfügen
Attribute mBtnAdd.VB_VarHelpID = -1
Private WithEvents mBtnAddCln As CommandButton ' 2. [ ++]  Klonen,      das ausgewählte Objekt Kopieren , Editieren, und am Ende der Liste anfügen
Attribute mBtnAddCln.VB_VarHelpID = -1
Private WithEvents mBtnIns    As CommandButton ' 3. [ i ]  Einfügen,    ein neues       Objekt Erstellen, Editieren, und oberhalb dem aktuell ausgewählten Objekt einfügen
Attribute mBtnIns.VB_VarHelpID = -1
Private WithEvents mBtnInsCln As CommandButton ' 4. [i++]  Kop+Einf,    das ausgewählte Objekt Kopieren , Editieren, und oberhalb dem aktuell ausgewählten Objekt einfügen
Attribute mBtnInsCln.VB_VarHelpID = -1
Private WithEvents mBtnEdit   As CommandButton ' 5. [ / ]  Edit,        das ausgewählte Objekt Editieren, einen Modalen Dialog mit den Eigenschaften des Objekts anzeigen, OK oder Abbrechen
Attribute mBtnEdit.VB_VarHelpID = -1
Private WithEvents mBtnDel    As CommandButton ' 6. [ - ]  Löschen,     das ausgewählte Objekt Löschen vorher nachfragen
Attribute mBtnDel.VB_VarHelpID = -1
Private WithEvents mBtnMovUp  As CommandButton ' 7. [ ^ ]  Nach Oben,   das ausgewählte Objekt um eine Stelle nach oben  schieben
Attribute mBtnMovUp.VB_VarHelpID = -1
Private WithEvents mBtnMovDwn As CommandButton ' 8. [ v ]  Nach Unten,  das ausgewählte Objekt um eine Stelle nach unten schieben
Attribute mBtnMovDwn.VB_VarHelpID = -1
Private WithEvents mBtnSrtUp  As CommandButton ' 9. [ |v]  Aufsteigend, alle Objekte aufsteigend Sortieren
Attribute mBtnSrtUp.VB_VarHelpID = -1
Private WithEvents mBtnSrtDwn As CommandButton '10. [ |^]  Absteigend,  alle Objekte absteigend  Sortieren
Attribute mBtnSrtDwn.VB_VarHelpID = -1
Private WithEvents mBtnSearch As CommandButton '11. [°\ ]  Suchen,      ein Objekt suchen, einen Suchdialog anzeigen, Ergebnisliste anzeigen
Attribute mBtnSearch.VB_VarHelpID = -1
Private WithEvents mTxtSearch As TextBox       '12. [Editbox]
Attribute mTxtSearch.VB_VarHelpID = -1
Private WithEvents mListBox   As ListBox
Attribute mListBox.VB_VarHelpID = -1
Private m_List                As CCollection
Private m_IsHashed            As Boolean
Public Event OnEdit(Obj_inout As Object)
Public Event OnFound(found As CCollection)

'höh man braucht die Info ob in der Col Keys oder Indices

'should we have the listbox also in here
'should we have the list also in here
Friend Sub New_(aList As CCollection, ListBox, BtnAdd, Optional BtnAddClone, Optional BtnInsert, Optional BtnInsertClone, Optional BtnEdit, Optional BtnDelete, Optional BtnMoveUp, Optional BtnMoveDown, Optional BtnSortUp, Optional BtnSortDown, Optional BtnSearch, Optional TxtSearch)
    Set m_List = aList: m_IsHashed = m_List.IsHashed: Set mListBox = ListBox
    If Not IsMissing(BtnAdd) Then Set mBtnAdd = BtnAdd
    If Not IsMissing(BtnAddClone) Then Set mBtnAddCln = BtnAddClone
    If Not IsMissing(BtnInsert) Then Set mBtnIns = BtnInsert
    If Not IsMissing(BtnInsertClone) Then Set mBtnInsCln = BtnInsertClone
    If Not IsMissing(BtnEdit) Then Set mBtnEdit = BtnEdit
    If Not IsMissing(BtnDelete) Then Set mBtnDel = BtnDelete
    If Not IsMissing(BtnMoveUp) Then Set mBtnMovUp = BtnMoveUp
    If Not IsMissing(BtnMoveDown) Then Set mBtnMovDwn = BtnMoveDown
    If Not IsMissing(BtnSortUp) Then Set mBtnSrtUp = BtnSortUp
    If Not IsMissing(BtnSortDown) Then Set mBtnSrtDwn = BtnSortDown
    If Not IsMissing(BtnSearch) Then Set mBtnSearch = BtnSearch
    If Not IsMissing(TxtSearch) Then Set mTxtSearch = TxtSearch
End Sub

'Copy to MNew.bas
'Public Function DBcrud(Col As CCollection, ByVal IsHashed As Boolean, ListBox, _
'                       BtnAdd As CommandButton, Optional BtnAddClone, Optional BtnInsert, Optional BtnInsertClone, Optional BtnEdit, _
'                       Optional BtnDelete, Optional BtnMoveUp, Optional BtnMoveDown, Optional BtnSortUp, Optional BtnSortDown, Optional BtnSearch, Optional TxtSearch) As DBcrud
'    Set DBcrud = New DBcrud: DBcrud.New_ Col, ListBox, BtnAdd, BtnAddClone, BtnInsert, BtnInsertClone, BtnEdit, BtnDelete, BtnMoveUp, BtnMoveDown, BtnSortUp, BtnSortDown, BtnSearch
'End Function
'
''How to use it?
''==============
''place all the buttons on the form, or maybe copy them from the Form FMain in this project, everything you need on a form is:
'Private WithEvents PersonCRUD As DBcrud
'Private Sub Form_Load()
'    Set PersonCRUD = MNew.DBcrud(MApp.Persons, False, Me.LstPersons, Me.BtnPersonAdd, Me.BtnPersonAddClone, Me.BtnPersonInsert, Me.BtnPersonInsertClone, Me.BtnPersonEdit, Me.BtnPersonDelete, Me.BtnPersonMoveUp, Me.BtnPersonMoveDown, Me.BtnPersonSortUp, Me.BtnPersonSortDwn, Me.BtnPersonSearch)
'End Sub
'Private Sub PersonCRUD_OnEdit(Obj_inout As Object)
'    Dim p As Person: Set p = IIf(Obj_inout Is Nothing, MNew.PersonDefault, Obj_inout)
'    If MNew.ModalDialog(FPerson, FPerson.BtnOK, FPerson.BtnCancel).ShowDialog(p, Me) = vbCancel Then
'        Set Obj_inout = Nothing
'        Exit Sub
'    End If
'    Set Obj_inout = p
'End Sub

Public Property Get List() As CCollection
    Set List = m_List
End Property
Public Property Set List(ByVal Values As CCollection)
    Set m_List = Values
End Property

Private Sub mBtnAdd_Click()
    Dim Obj As Object
    RaiseEvent OnEdit(Obj)
    If Obj Is Nothing Then Exit Sub
    m_List.Add Obj
    mListBox.AddItem Obj.ToStr
    mListBox.ItemData(mListBox.ListCount - 1) = Obj.Ptr
End Sub
Private Sub mBtnAddCln_Click()
    Dim li As Long: li = mListBox.ListIndex
    If li < 0 Then MsgBox "Select object first": Exit Sub
    Dim Obj As Object: Set Obj = m_List.Item(li + 1)
    Set Obj = Obj.Clone
    'should we show the edit dlg here?
    'yes because with it the user gets the chance to cancel this command
    RaiseEvent OnEdit(Obj)
    If Obj Is Nothing Then Exit Sub
    m_List.Add Obj
    mListBox.AddItem Obj.ToStr
    mListBox.ItemData(mListBox.ListCount - 1) = Obj.Ptr
End Sub

Private Sub mBtnIns_Click()
    Dim Obj As Object
    RaiseEvent OnEdit(Obj)
    If Obj Is Nothing Then Exit Sub
    Dim li As Long: li = mListBox.ListIndex
    m_List.Insert li + 1, Obj
    mListBox.AddItem Obj.ToStr, li
    mListBox.ItemData(li) = Obj.Ptr
    mListBox.ListIndex = li
End Sub
Private Sub mBtnInsCln_Click()
    Dim li As Long: li = mListBox.ListIndex
    If li < 0 Then MsgBox "Select object first": Exit Sub
    Dim Obj As Object: Set Obj = m_List.Item(li + 1)
    Set Obj = Obj.Clone
    RaiseEvent OnEdit(Obj)
    If Obj Is Nothing Then Exit Sub
    m_List.Insert li + 1, Obj
    mListBox.AddItem Obj.ToStr, li
    mListBox.ItemData(li) = Obj.Ptr
    mListBox.ListIndex = li
End Sub

Private Sub mBtnEdit_Click()
    Dim i As Long: i = -1
    Dim Obj As Object: Set Obj = m_List.ObjectFromListCtrl(mListBox, i)
    If Obj Is Nothing Then MsgBox "First select an object from the list!": Exit Sub
    RaiseEvent OnEdit(Obj)
    If Obj Is Nothing Then Exit Sub
    mListBox.List(i) = Obj.ToStr
End Sub

Private Sub mBtnDel_Click()
    Dim li As Long: li = mListBox.ListIndex
    Dim i As Long: i = -1
    Dim Obj As Object: Set Obj = m_List.ObjectFromListCtrl(mListBox, i)
    If Obj Is Nothing Then MsgBox "First select an object from the list!": Exit Sub
    Dim mr As VbMsgBoxResult: mr = MsgBox("Do you really want to delete the object!" & vbCrLf & Obj.ToStr, vbOKCancel)
    If mr = VbMsgBoxResult.vbCancel Then Exit Sub
    m_List.Remove Obj.Key
    m_List.ToListBox mListBox, , True
    If mListBox.ListCount Then mListBox.ListIndex = li
End Sub

Private Sub mBtnMovUp_Click()
    Dim li As Long: li = mListBox.ListIndex
    If li < 0 Then MsgBox "Select object first": Exit Sub
    If li = 0 Then Exit Sub 'already the topmost object
    m_List.MoveUp li + 1
    LBCB_MoveUp mListBox, li
    mListBox.ListIndex = li - 1
End Sub
Private Sub mBtnMovDwn_Click()
    Dim li As Long: li = mListBox.ListIndex
    If li < 0 Then MsgBox "Select object first": Exit Sub
    If (mListBox.ListCount - 1) = li Then Exit Sub 'already the lowest object
    m_List.MoveDown li + 1
    LBCB_MoveDown mListBox, li
    mListBox.ListIndex = li + 1
End Sub

Private Sub LBCB_MoveUp(LBCB, ByVal i As Long)
    LBCB_Swap LBCB, i, i - 1
End Sub
Private Sub LBCB_MoveDown(LBCB, ByVal i As Long)
    LBCB_Swap LBCB, i, i + 1
End Sub
Private Sub LBCB_Swap(LBCB, ByVal i1 As Long, ByVal i2 As Long)
    With LBCB 'As ListBox Or As ComboBox
        Dim s As String: s = .List(i1)
        Dim id As Long: id = .ItemData(i1)
        .List(i1) = .List(i2)
        .ItemData(i1) = .ItemData(i2)
        .List(i2) = s
        .ItemData(i2) = id
    End With
End Sub

Private Sub mBtnSrtUp_Click()
    m_List.Sort
    m_List.ToListBox mListBox, , True
End Sub
Private Sub mBtnSrtDwn_Click()
    m_List.Sort True
    m_List.ToListBox mListBox, , True
End Sub

Private Sub mBtnSearch_Click()
    Dim s As String: s = mTxtSearch.Text
    If Len(s) = 0 Then Exit Sub
    Dim found As CCollection: Set found = m_List.Find(s)
    If found Is Nothing Then Exit Sub
    If found.Count = 0 Then Exit Sub
    RaiseEvent OnFound(found)
End Sub

Private Sub mListBox_Click()
    '
End Sub

Private Sub mListBox_DblClick()
    mBtnEdit.Value = True
End Sub

Private Sub mListBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case KeyCodeConstants.vbKeyReturn:   mBtnEdit.Value = True
    Case KeyCodeConstants.vbKeyAdd:      mBtnAdd.Value = True
    Case KeyCodeConstants.vbKeyMultiply: mBtnAddCln.Value = True
    Case KeyCodeConstants.vbKeySubtract: mBtnDel.Value = True
    Case KeyCodeConstants.vbKeyDivide:   mBtnEdit.Value = True
    Case KeyCodeConstants.vbKeyI:        mBtnIns.Value = True
    End Select
End Sub

Private Sub mListBox_KeyPress(KeyAscii As Integer)
    '
End Sub

Private Sub mListBox_KeyUp(KeyCode As Integer, Shift As Integer)
    '
End Sub

Private Sub mTxtSearch_Change()
    '
End Sub

Private Sub mTxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    '
End Sub

Private Sub mTxtSearch_KeyPress(KeyAscii As Integer)
    '
End Sub

Private Sub mTxtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyReturn Then mBtnSearch.Value = True '_Click
End Sub
