VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModalDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Result     As VbMsgBoxResult
Private m_Object     As Object
Private m_Dialog     As Object
Private m_DlgFrm     As Form
Private m_IsReadOnly As Boolean
Private WithEvents mBtnOK     As CommandButton
Attribute mBtnOK.VB_VarHelpID = -1
Private WithEvents mBtnCancel As CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1

Private Sub Class_Initialize()
    m_Result = VbMsgBoxResult.vbCancel
End Sub

Friend Sub New_(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton, Optional IsReadOnly As Boolean = False)
    Set m_Dialog = Dialog: Set m_DlgFrm = Dialog: Set mBtnOK = BtnOK: Set mBtnCancel = BtnCancel: m_IsReadOnly = IsReadOnly
End Sub

'Copy to MNew.bas:
'Public Function ModalDialog(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton, Optional ByVal IsReadOnly As Boolean = False) As ModalDialog
'    Set ModalDialog = New ModalDialog: ModalDialog.New_ Dialog, BtnOK, BtnCancel, IsReadOnly
'End Function

Public Function ShowDialog(Obj As Object, Owner As Form) As VbMsgBoxResult
    If m_IsReadOnly Then Set m_Object = Obj Else Set m_Object = Obj.Clone
    m_Dialog.UpdateView m_Object, m_IsReadOnly
    m_DlgFrm.Show vbModal, Owner
    ShowDialog = m_Result
    If ShowDialog = vbCancel Then Exit Function
    If Not m_IsReadOnly Then Obj.NewC m_Object
End Function
'Q: could this class also be used as an Interface for a dialog,
'so the extra interface class could maybe be obsolete?
'Showdialog is the only function in the class ModalDialog
'But the window-dlg only must implement UpdateView and UpdateData
'no the function ShowDialog is the game-stopper
'Q: could the function ShowDialog be Friend?
'A: no because unfortunately the function ShowDialog must be Public otherwise it will not be found from a calling object
'Public Sub UpdateView(Obj, Optional ByVal IsReadOnly As Boolean = False): End Sub
'Public Function UpdateData(Obj) As Boolean: End Function

Private Sub mBtnOK_Click()
    If Not m_IsReadOnly Then
        If Not m_Dialog.UpdateData(m_Object) Then Exit Sub
    End If
    m_Result = VbMsgBoxResult.vbOK
    Unload m_Dialog
End Sub

Private Sub mBtnCancel_Click()
    Unload m_Dialog
End Sub

'Example for using this class in a ButtonClick on a Form:
'Private Sub BtnCityAdd_Click()
'    Dim City As City: Set City = MNew.City("Musterstadt", "00000")
'    If MNew.ModalDialog(FCity, FCity.BtnOK, FCity.BtnCancel).ShowDialog(City, Me) = vbCancel Then Exit Sub
'    MData.Cities_AddLB LstCities, City
'End Sub
'in the Form we need:
'Option Explicit
'Private m_Person As Person
'
'Private Sub Form_Load()
'    MApp.Cities.ToListCtrl CmbCity, , True
'End Sub
'
'Public Sub UpdateView(obj, Optional ByVal IsReadOnly As Boolean = False)
'    Set m_Person = obj
'    If m_Person Is Nothing Then MsgBox "The Person does not exist": Exit Sub
'    TxtName.Text = m_Person.Name
'    TxtBirthDay.Text = m_Person.BirthDay
'    If m_Person.City Is Nothing Then MsgBox "The City does not exist": Exit Sub
'    CmbCity.ListIndex = MApp.Cities.IndexFromObject(m_Person.City)
'    CmbCity.Text = m_Person.City.ToStr
'    If IsReadOnly Then
'        Me.TxtName.Enabled = False
'        Me.TxtBirthDay.Enabled = False
'        Me.CmbCity.Enabled = False
'    End If
'End Sub
'
'Public Function UpdateData(obj) As Boolean
'Try: On Error GoTo Catch
'    Dim bIsOK As Boolean
'    Dim bd As Date:     bd = m_Person.BirthDay: TxtBirthDay.Text = MString.Date_TryParseValidate(TxtBirthDay.Text, "Birthday", "", bIsOK, bd): If Not bIsOK Then Exit Function
'    Dim ic As Long: ic = -1
'    Dim ct As City: Set ct = MApp.Cities.ObjectFromListCtrl(CmbCity, ic)
'    Dim ii As Long:     ii = m_Person.Index
'    Dim nm As String:   nm = TxtName.Text
'    m_Person.New_ bd, ct, ii, nm
'    UpdateData = True
'Catch:
'End Function
'
'Private Sub TxtBirthDay_Validate(Cancel As Boolean)
'    Dim bd As Date: bd = m_Person.BirthDay
'    Dim bIsOK As Boolean: TxtBirthDay.Text = MString.Date_TryParseValidate(TxtBirthDay.Text, "Birthday", "", bIsOK, bd)
'    If bIsOK Then m_Person.SetParams BirthDay:=bd
'    Cancel = Not bIsOK
'End Sub


