VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'om 2025-11-30 lines: 410; 2025-12-04 lines: 464;
'Class CCollection : replaces and extends the vb intrinsic-Collection with new additional functions
'works as a wrapper class for your Collections, is fully compatible to your existing Code
'the VB-Collection is 1-based and so is CCollection as we do not break with compatibility
'Implements Collection 'nope
Private m_Name      As String
Private m_Col       As VBA.Collection 'Of Object/Variant
Private m_IsHashed  As Boolean
Private m_Version   As Variant 'Decimal 'Incremented in every Add, Insert or Delete
Private m_LowB      As Long
Private m_SortDirUp As Boolean

Private Sub Class_Initialize()
    Set m_Col = New Collection
End Sub

Friend Sub New_(IsHashed As Boolean, Optional col As Collection = Nothing, Optional Name As String = "", Optional OptionBaseLBound As Long = 1)
    'works as a wrapper and extends existing VBA.Collection.
    'The default LBound aka Option Base of a VBA.Collection is 1, and so is here, but now you are free to define other LBound as well
    m_Name = Name: m_IsHashed = IsHashed: If Not col Is Nothing Then Set m_Col = col
    m_LowB = OptionBaseLBound - 1
End Sub

'Copy to MNew
'Public Function CCollection(ByVal IsHashed As Boolean, Optional Col As Collection = Nothing, Optional Name As String = "", Optional OptionBaseLBound As Long = 1) As CCollection
'    Set CCollection = New CCollection: CCollection.New_ IsHashed, Col, Name, OptionBaseLBound
'End Function

Public Sub NewC(other As CCollection)
    Dim i As Long
    With other
        m_Name = .Name
        m_IsHashed = .IsHashed
        m_LowB = .LLBound
        For i = .LLBound To .UUBound
            m_Col.Add .Item(i + m_LowB)
        Next
    End With
End Sub

'Optional includingObjects As Boolean = False
Public Function Clone() As CCollection
    Set Clone = New CCollection: Clone.NewC Me
End Function

Public Function Compare(other As CCollection) As Long
    Compare = StrComp(Me.Name, other.Name, vbBinaryCompare)
End Function

Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(ByVal Value As String)
    m_Name = Value
End Property

Public Function Find(Text As String) As CCollection
    Set Find = MNew.CCollection(Me.IsHashed)
    Dim Obj As Object
    For Each Obj In m_Col
        If Obj.Find(Text) Then Find.Add Obj
    Next
End Function

Public Function Equals(other As CCollection) As Boolean
    If Not other.Count = Me.Count Then Exit Function
    Dim i As Long
    For i = 1 To m_Col.Count
        If m_Col.Item(i) <> other.Item(i) Then Exit Function
    Next
    Equals = True
End Function

Public Property Get Ptr()
    Ptr = ObjPtr(Me)
End Property

Public Property Get Key() As String
    If Len(m_Name) Then Key = Me.Name Else Key = Me.Ptr
End Property

Public Property Get DataType() As VbVarType
    If m_Col.Count = 0 Then Exit Property
    DataType = VarType(m_Col.Item(1))
End Property

Private Function DataType_ToStr(ByVal vt As VbVarType) As String
    Dim s As String
    Select Case vt
    Case VbVarType.vbArray:           s = "Array"
    Case VbVarType.vbBoolean:         s = "Boolean"
    Case VbVarType.vbByte:            s = "Byte"
    Case VbVarType.vbCurrency:        s = "Currency"
    Case VbVarType.vbDataObject:      s = "DataObject"
    Case VbVarType.vbDate:            s = "Date"
    Case VbVarType.vbDecimal:         s = "Decimal"
    Case VbVarType.vbDouble:          s = "Double"
    Case VbVarType.vbEmpty:           s = "Empty"
    Case VbVarType.vbError:           s = "Error"
    Case VbVarType.vbInteger:         s = "Integer"
    Case VbVarType.vbLong:            s = "Long"
    Case VbVarType.vbNull:            s = "Null"
    Case VbVarType.vbObject:  Dim Obj As Object: Set Obj = m_Col.Item(1)
                                      s = "Object(" & TypeName(Obj) & ")"
    Case VbVarType.vbSingle:          s = "Single"
    Case VbVarType.vbString:          s = "String"
    Case VbVarType.vbUserDefinedType: s = "UserDefinedType"
    Case VbVarType.vbVariant:         s = "Variant"
    End Select
    DataType_ToStr = s
End Function

Public Function Add(Item, Optional Key, Optional Before, Optional After)
    If IsObject(Item) Then
        Dim Obj As Object: Set Obj = Item
        Set Add = Obj
        If m_IsHashed Then
            If IsMissing(Key) Then Key = Obj.Key
        End If
    End If
    If IsMissing(Key) Then
        If IsMissing(Before) Then
            If IsMissing(After) Then
                m_Col.Add Item
            Else
                m_Col.Add Item, , , After
            End If
        Else
            m_Col.Add Item, , Before 'before togehter with after makes no sense
        End If
    Else
        If IsMissing(Before) Then
            If IsMissing(After) Then
                m_Col.Add Item, Key
            Else
                m_Col.Add Item, Key, , After
            End If
        Else
            m_Col.Add Item, Key, Before
        End If
    End If
    IncVersion
End Function

Public Function AddKey(Obj As Object) As Object
    ' the object needs to have a Public Function/PropertyGet Key As String
    Set AddKey = Obj: m_Col.Add Obj, Obj.Key
    IncVersion
End Function

Public Function AddOrGet(Obj As Object) As Object
    Dim Key As String: Key = Obj.Key ' the object needs to have a Public Function/PropertyGet Key As String
    If Me.Contains(Key) Then
        Set AddOrGet = m_Col.Item(Key)
    Else
        Set AddOrGet = Obj
        m_Col.Add Obj, Key
        m_Version = m_Version + CDec(1)
    End If
End Function

Public Function TryAddObject(Obj As Object, ByVal Key As String) As Boolean
Try: On Error GoTo Catch
    m_Col.Add Obj, Key
    TryAddObject = True
    IncVersion
Catch: On Error GoTo 0
End Function

Public Sub ChangeKey(OldIndexKey, NewKey As String)
    'for this function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original version of the function "CollectionChangeKey" here: https://vb-tec.de/collctns.htm
    Dim Value As Variant
    If IsObject(m_Col(OldIndexKey)) Then
        Set Value = m_Col.Item(OldIndexKey)
    Else
        Value = m_Col.Item(OldIndexKey)
    End If
    m_Col.Add Value, NewKey  'first add
    m_Col.Remove OldIndexKey 'then remove
    IncVersion
End Sub

Public Sub Clear()
    Class_Initialize
    IncVersion
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_Col.[_NewEnum]
End Function

Public Function Contains(ByVal Key As String) As Boolean
    'for this function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original version of the function "IsInCollection" here: https://vb-tec.de/collctns.htm
    On Error Resume Next
    '"Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(m_Col(Key)) Then: 'DoNothing
    Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

Private Function GetIndex(ByVal i As Long) As Long
    GetIndex = i - m_LowB
End Function

Public Property Get LLBound() As Long
    LLBound = 1 + m_LowB
End Property

Public Property Get UUBound() As Long
    'gibt den Upper Bound, die oberste Schranke, den größten Index des internen Arrays zurück
    UUBound = m_Col.Count + m_LowB
End Property

'Public Property Get UpperBound() As Long
'    UpperBound = Me.Count - m_OptBase
'End Property

Public Function IndexFromObject(Obj As Object) As Long
    Dim i As Long, V, o As Object
    For Each V In m_Col
        Set o = V
        If o.Key = Obj.Key Then
            IndexFromObject = i
            Exit Function
        End If
        i = i + 1
    Next
End Function

Public Property Get IsHashed() As Boolean
    IsHashed = m_IsHashed
End Property

Public Sub Insert(ByVal Index As Long, ValueOrObj)
    'insert always means "Befor" the Index
    'you can insert before the first element to get a new first element
    'if you want to insert after the last element just use Add
    Index = GetIndex(Index)
    Me.Add ValueOrObj, , Index
    'm_Col.Add ValueOrObj, , Index
    'm_Version = m_Version + CDec(1)
End Sub

Public Property Get Item(IndexOrKey)
    If VarType(IndexOrKey) = vbString Then
        'IndexOrKey is a Key!
        If IsObject(m_Col.Item(IndexOrKey)) Then
            Set Item = m_Col.Item(IndexOrKey)
        Else
            Item = m_Col.Item(IndexOrKey)
        End If
    Else
        'IndexOrKey is a numeric Index!
        Dim i As Long: i = GetIndex(IndexOrKey)
        If IsObject(m_Col.Item(i)) Then
            Set Item = m_Col.Item(i)
        Else
            Item = m_Col.Item(i)
        End If
    End If
End Property

Public Property Let Item(IndexOrKey, ValueOrObj)
    Dim vtik As VbVarType: vtik = VarType(IndexOrKey)
    If vtik = vbString Then 'its a key
        m_Col.Remove IndexOrKey
        m_Col.Add ValueOrObj, IndexOrKey
    ElseIf vtik = vbLong Or vtik = vbInteger Then
        'IndexOrKey is a numeric Index!
        Dim i As Long: i = GetIndex(IndexOrKey)
        m_Col.Remove i
        m_Col.Add ValueOrObj, , i
    End If
    IncVersion
End Property

Public Property Set Item(IndexOrKey, ValueOrObj As Object)
    Dim vtik As VbVarType: vtik = VarType(IndexOrKey)
    If vtik = vbString Then 'its a key
        m_Col.Remove IndexOrKey
        m_Col.Add ValueOrObj, IndexOrKey
    ElseIf vtik = vbLong Or vtik = vbInteger Then
        'IndexOrKey is a numeric Index!
        Dim i As Long: i = GetIndex(IndexOrKey)
        m_Col.Remove i
        m_Col.Add ValueOrObj, , i
    End If
    IncVersion
End Property

Public Sub MoveUp(ByVal i As Long)
    If m_IsHashed Then MoveUpKey i: Exit Sub
    Dim c As Long: c = m_Col.Count
    If i <= 1 Or c < i Then Exit Sub
    SwapItems i, i - 1
    IncVersion
End Sub

Public Sub MoveUpKey(ByVal i As Long)
    Dim c As Long: c = m_Col.Count
    If i <= 1 Or c < i Then Exit Sub
    Dim i1 As Long: i1 = i
    Dim i2 As Long: i2 = i - 1
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    Dim Key1 As String: Key1 = Obj1.Key
    Dim Key2 As String: Key2 = Obj2.Key
    SwapItemsKey i1, Key1, i2, Key2
    IncVersion
End Sub

Public Sub MoveDown(ByVal i As Long)
    If m_IsHashed Then MoveDownKey i: Exit Sub
    Dim c As Long: c = m_Col.Count
    If i < 1 Or c <= i Then Exit Sub
    SwapItems i, i + 1
    m_Version = m_Version + CDec(1)
End Sub

Public Sub MoveDownKey(ByVal i As Long)
    Dim c As Long: c = m_Col.Count
    If i < 1 Or c <= i Then Exit Sub
    Dim i1 As Long: i1 = i
    Dim i2 As Long: i2 = i + 1
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    Dim Key1 As String: Key1 = Obj1.Key
    Dim Key2 As String: Key2 = Obj2.Key
    SwapItemsKey i1, Key1, i2, Key2
    IncVersion
End Sub

Public Function Remove(IndexOrKey)
    Set Remove = Me.Item(IndexOrKey)
    m_Col.Remove IndexOrKey
    m_Version = m_Version + CDec(1)
End Function

Public Sub SwapItems(ByVal i1 As Long, ByVal i2 As Long)
    Dim c As Long: c = m_Col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or c <= i1 Then Exit Sub
    If i2 <= 0 Or c < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim Obj1: If IsObject(m_Col.Item(i1)) Then Set Obj1 = m_Col.Item(i1) Else Obj1 = m_Col.Item(i1)
    Dim Obj2: If IsObject(m_Col.Item(i2)) Then Set Obj2 = m_Col.Item(i2) Else Obj2 = m_Col.Item(i2)
    m_Col.Remove i1:
    m_Col.Add Obj2, , i1
    m_Col.Remove i2
    If i2 < c Then
        m_Col.Add Obj1, , i2
    Else
        m_Col.Add Obj1
    End If
    IncVersion
End Sub

Public Sub SwapItemsKey(ByVal i1 As Long, ByVal Key1 As String, _
                        ByVal i2 As Long, ByVal Key2 As String)
    Dim c As Long: c = m_Col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or c <= i1 Then Exit Sub
    If i2 <= 0 Or c < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim Obj1: If IsObject(m_Col.Item(Key1)) Then Set Obj1 = m_Col.Item(Key1) Else Obj1 = m_Col.Item(Key1)
    Dim Obj2: If IsObject(m_Col.Item(Key2)) Then Set Obj2 = m_Col.Item(Key2) Else Obj2 = m_Col.Item(Key2)
    m_Col.Remove Key2: m_Col.Add Obj2, Key2, i1
    m_Col.Remove Key1: m_Col.Add Obj1, Key1, i1
    IncVersion
End Sub

#If VBA7 Then
    Public Sub ToListBox(aLB As MSForms.ListBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
#Else
    Public Sub ToListBox(aLB As ListBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
#End If
    ToListCtrl aLB, addEmptyLineFirst, doPtrToItemData
End Sub

#If VBA7 Then
    Public Sub ToComboBox(aCB As MSForms.ComboBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
#Else
    Public Sub ToComboBox(aCB As ComboBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
#End If
    ToListCtrl aCB, addEmptyLineFirst, doPtrToItemData
End Sub

Public Sub ToListCtrl(ComboBoxOrListBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    Dim i As Long, c As Long: c = m_Col.Count: If c = 0 Then Exit Sub
    Dim vt As VbVarType: vt = VarType(m_Col.Item(1))
    Dim V, Obj As Object
    With ComboBoxOrListBox
        If .ListCount Then .Clear
        If addEmptyLineFirst Then .AddItem vbNullString
        Select Case vt
        Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal, vbString
            For i = 1 To c
                .AddItem m_Col.Item(i)
            Next
        Case vbObject
            For i = 1 To c
                Set Obj = m_Col.Item(i)
                .AddItem Obj.ToStr ' the object needs to have a Public Function ToStr As String
                If doPtrToItemData Then .ItemData(i - 1) = Obj.Ptr ': Debug.Print Obj.Ptr ' and a Public Function Ptr As LongPtr
                ' ItemData can only be of type Long no String
            Next
        End Select
    End With
End Sub

Public Property Get ObjectFromListCtrl(ComboBoxOrListBox, ByRef i_out As Long) As Object
    Dim li As Long: li = ComboBoxOrListBox.ListIndex
    If i_out < 0 Then i_out = li
    If i_out < 0 Then Exit Property
    Dim Key As String: Key = ComboBoxOrListBox.ItemData(i_out)
    If Contains(Key) Then Set ObjectFromListCtrl = m_Col.Item(Key)
End Property

Public Function ToStr(Optional ByVal isLong As Boolean = False) As String
    ToStr = m_Name
    If Not isLong Then Exit Function
    ToStr = TypeName(Me) & "(Of " & DataType_ToStr(Me.DataType) & ")" & " " & m_Name
End Function

Public Function Data_ToStr() As String 'Col As Collection) As String
    Dim s As String, V, Obj As Object
    For Each V In m_Col
        If IsObject(V) Then
            Set Obj = V
            s = s & Obj.ToStr & vbCrLf
        Else
            s = s & CStr(V) & vbCrLf
        End If
    Next
    Data_ToStr = s
End Function

Public Property Get Version() As Variant 'As Decimal
    Version = m_Version
End Property

Private Sub IncVersion()
    m_Version = m_Version + CDec(1)
End Sub
' Sort:
Public Sub Sort(Optional SortDirectionDown As Boolean = False)
    m_SortDirUp = Not SortDirectionDown
    'Set m_Col = Col
    Dim c As Long: c = m_Col.Count
    'If c = 0 Then: Set m_Col = Nothing: Exit Sub
    If c = 0 Then Exit Sub
    Dim vt As VbVarType: vt = VarType(m_Col.Item(1))
    Select Case vt
    Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal
        QuickSortVar 1, c
    Case vbString
        QuickSortStr 1, c
    Case vbObject
        QuickSortObj 1, c
    End Select
    'Set m_Col = Nothing
End Sub
    
' The recursive data-independent QuickSort for primitive data-variables
Private Sub QuickSortVar(ByVal i1 As Long, ByVal i2 As Long)
    Dim t As Long
    If i2 > i1 Then
        t = DivideVar(i1, i2)
        QuickSortVar i1, t - 1
        QuickSortVar t + 1, i2
    End If
End Sub

Private Function DivideVar(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareVar(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareVar(p, j) < 0))
        If i < j Then SwapVar i, j
    Loop While (i < j)
    SwapVar i, p
    DivideVar = i
End Function

Private Function CompareVar(ByVal i1 As Long, ByVal i2 As Long) As Variant
    If m_SortDirUp Then
        CompareVar = m_Col.Item(i1) - m_Col.Item(i2)
    Else
        'Dim t As Long: t = i1: i1 = i2: i2 = t
        CompareVar = m_Col.Item(i2) - m_Col.Item(i1)
    End If
End Function

Private Sub SwapVar(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Var1: Var1 = m_Col.Item(i1)
    Dim Var2: Var2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Var2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Var1, , i2 Else m_Col.Add Var1
End Sub

' The recursive data-independent QuickSort for strings
Private Sub QuickSortStr(ByVal i1 As Long, ByVal i2 As Long)
    Dim t As Long
    If i1 < i2 Then
        t = DivideStr(i1, i2)
        QuickSortStr i1, t - 1
        QuickSortStr t + 1, i2
    End If
End Sub

Private Function DivideStr(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareStr(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareStr(p, j) < 0))
        If i < j Then SwapStr i, j
    Loop While (i < j)
    SwapStr i, p
    DivideStr = i
End Function

Private Function CompareStr(ByVal i1 As Long, ByVal i2 As Long)
    If m_SortDirUp Then
        CompareStr = StrComp(m_Col.Item(i1), m_Col.Item(i2))
    Else
        CompareStr = StrComp(m_Col.Item(i2), m_Col.Item(i1))
    End If
    'Dim Str1 As String: Str1 = m_col.Item(i1)
    'Dim Str2 As String: Str2 = m_col.Item(i2)
    'CompareStr = StrComp(Str1, Str2)
End Function

Private Sub SwapStr(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Str1 As String: Str1 = m_Col.Item(i1)
    Dim Str2 As String: Str2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Str2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Str1, , i2 Else m_Col.Add Str1
End Sub

' The recursive data-independent QuickSort for objects
Private Sub QuickSortObj(ByVal i1 As Long, ByVal i2 As Long)
    Dim t As Long
    If i2 > i1 Then
        t = DivideObj(i1, i2)
        QuickSortObj i1, t - 1
        QuickSortObj t + 1, i2
    End If
End Sub

Private Function DivideObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareObj(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareObj(p, j) < 0))
        If i < j Then SwapObj i, j
    Loop While (i < j)
    SwapObj i, p
    DivideObj = i
End Function

Private Function CompareObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    If m_SortDirUp Then
        CompareObj = Obj1.Compare(Obj2)
    Else
        CompareObj = Obj2.Compare(Obj1)
    End If
End Function

Private Sub SwapObj(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    If m_IsHashed Then
        
        Dim Key1 As String: Key1 = Obj1.Key
        Dim Key2 As String: Key2 = Obj2.Key
        Me.SwapItemsKey i1, Key1, i2, Key2
        
        m_Col.Remove Key2
        m_Col.Add Obj2, Key2, i1
        
        m_Col.Remove Key1
        If i2 < c Then m_Col.Add Obj1, Key1, i2 Else m_Col.Add Obj1, Key1
    Else
        m_Col.Remove i1: m_Col.Add Obj2, , i1:   m_Col.Remove i2
        If i2 < c Then m_Col.Add Obj1, , i2 Else m_Col.Add Obj1
    End If
    
End Sub

'Public Sub SwapItemsKey(ByVal i1 As Long, ByVal Key1 As String, _
'                        ByVal i2 As Long, ByVal Key2 As String)
'    Dim c As Long: c = m_Col.Count
'    If c = 0 Then Exit Sub
'    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
'    If i1 <= 0 Or c <= i1 Then Exit Sub
'    If i2 <= 0 Or c < i2 Then Exit Sub
'    If i1 = i2 Then Exit Sub
'    Dim Obj1: If IsObject(m_Col.Item(Key1)) Then Set Obj1 = m_Col.Item(Key1) Else Obj1 = m_Col.Item(Key1)
'    Dim Obj2: If IsObject(m_Col.Item(Key2)) Then Set Obj2 = m_Col.Item(Key2) Else Obj2 = m_Col.Item(Key2)
'    m_Col.Remove Key2: m_Col.Add Obj2, Key2, i1
'    m_Col.Remove Key1: m_Col.Add Obj1, Key1, i1
'    m_Version = m_Version + CDec(1)
'End Sub


