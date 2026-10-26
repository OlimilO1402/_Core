VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Version"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'     Member von: System ' Zeilen: 202
'Instancing: 6 - GlobalMultiUse
Option Explicit 'Ok wow is fertig oder?
'Implements ICloneable
'Implements IComparable
Private m_Major As Long
Private m_Minor As Long
Private m_Build As Long
Private m_Revis As Long

Private Sub Class_Initialize()
    'm_Major = 0: m_Minor = 0
    m_Build = -1: m_Revis = -1
End Sub

'Public Sub New()
'Public Sub New(ByVal major As Integer, ByVal minor As Integer)
'Public Sub New(ByVal major As Integer, ByVal minor As Integer, ByVal build As Integer)
'Public Sub New(ByVal major As Integer, ByVal minor As Integer, ByVal build As Integer, ByVal revision As Integer)
'Public Sub New(ByVal version As String)
Public Sub New_(ByVal Major As Long, ByVal Minor As Long, Optional ByVal Build, Optional ByVal Revision) 'As System.Version
    m_Major = Major: m_Minor = Minor: If Not IsMissing(Build) Then m_Build = Build: If Not IsMissing(Revision) Then m_Revis = Revision
End Sub

Public Sub Parse(Value As String)
    Dim StrVals() As String, StrVal As String, n As Long
    StrVals = Split(Value, ".")
    n = UBound(StrVals) - LBound(StrVals) + 1
    If n > 0 Then
        StrVal = StrVals(0)
        If IsNumeric(StrVal) Then m_Major = CLng(Trim(StrVal))
        If n > 1 Then
            StrVal = StrVals(1)
            If IsNumeric(StrVal) Then m_Minor = CLng(Trim(StrVal))
            If n > 2 Then
                StrVal = StrVals(2)
                If IsNumeric(StrVal) Then m_Build = CLng(Trim(StrVal))
                If n > 3 Then
                    StrVal = StrVals(3)
                    If IsNumeric(StrVal) Then m_Revis = CLng(Trim(StrVal))
                End If
            End If
        End If
    End If
End Sub

'Public Function Version(Optional ByVal Major As Long, Optional ByVal Minor As Long, Optional ByVal Build, Optional ByVal Revision) As Version
'    Set Version = New Version: Version.New_ Major, Minor, Build, Revision
'End Function
'Public Function VersionV(ByVal StrVersion As String) As Version
'    set VersionV = New Version: Version.NewV StrVersion
'End Function

'Private Function ICloneable_Clone() As Object
'  Set ICloneable_Clone = Clone
'End Function
'Public Overridable NotOverridable Function Clone() As Object
'Public Function Clone() As Object
'  Set Clone = Me
'End Function

'Rückgabewerte:
' Rückgabewert    Beschreibung
' Kleiner als 0   Dies ist die Instanz einer früheren Version als OVersion.
'             0   Diese Instanz entspricht OVersion.
'  Größer als 0   Dies ist die Instanz einer späteren Version als OVersion.
'                 – oder –   version ist null.
'Private Function IComparable_CompareTo(ByVal obj As Object) As cInteger
'  Set IComparable_CompareTo = CompareTo(obj)
'End Function
'Public Overridable NotOverridable Function CompareTo(ByVal version As Object) As Integer
'Public Function CompareTo(ByVal OVersion As Object) As Long 'cInteger
'  Set CompareTo = New cInteger
'  If TypeOf OVersion Is System.Version Then
'    If Major.CompareTo(OVersion.Major) > 0 Then
'      CompareTo = 1
'    Else
'      If Major.CompareTo(OVersion.Major) = 0 Then
'        If Minor.CompareTo(OVersion.Minor) > 0 Then
'          CompareTo = 1
'        Else
'          If Minor.CompareTo(OVersion.Minor) = 0 Then
'            If Build.CompareTo(OVersion.Build) > 0 Then
'              CompareTo = 1
'            Else
'              If Build.CompareTo(OVersion.Build) = 0 Then
'                If Revision.CompareTo(OVersion.Revision) > 0 Then
'                  CompareTo = 1
'                Else
'                  If Revision.CompareTo(OVersion.Revision) = 0 Then
'                    CompareTo = 0 'die beiden sind gleich
'                  Else
'                    CompareTo = -1
'                  End If
'                End If
'              Else
'                CompareTo = -1
'              End If
'            End If
'          Else
'            CompareTo = -1
'          End If
'        End If
'      Else
'        CompareTo = -1
'      End If
'    End If
'  End If
'End Function

Public Function Compare(other As Version) As Long
    'if this version is smaller than other version then return -1
    'if this version is greater then other version then return 1
    'both are equal return 0
    'if this.version < other.version then return -1
    'if other.version < this.version then return 1
    'else return 0
    
    'check Major
    Compare = m_Major < other.Major:      If Compare Then Exit Function 'return -1
    Compare = Abs(m_Major < other.Major): If Compare Then Exit Function 'return  1
    
    'check Minor
    Compare = m_Minor < other.Minor:      If Compare Then Exit Function 'return -1
    Compare = Abs(m_Minor < other.Major): If Compare Then Exit Function 'return  1
    
    'check Build
    Compare = m_Build < other.Build:      If Compare Then Exit Function 'return -1
    Compare = Abs(m_Build < other.Build): If Compare Then Exit Function 'return  1
    
    'check Revision
    Compare = m_Revis < other.Revision:      If Compare Then Exit Function 'return -1
    Compare = Abs(m_Revis < other.Revision): If Compare Then Exit Function 'return  1
    'return 0
End Function

Public Function Equals(other As Version) As Boolean
    Equals = m_Major = other.Major:    If Not Equals Then Exit Function
    Equals = m_Minor = other.Minor:    If Not Equals Then Exit Function
    Equals = m_Build = other.Build:    If Not Equals Then Exit Function
    Equals = m_Revis = other.Revision ': If Not Equals Then Exit Function
End Function

'Public Overrides Function Equals(ByVal obj As Object) As Boolean
'Public Function Equals(ByVal obj As Object) As cBoolean
'  Set Equals = New cBoolean
'End Function

'Public Overrides Function GetHashCode() As Integer
'Public Function GetHashCode() As cInteger
'  '
'End Function

'Public Shared Function op_Equality(ByVal v1 As System.Version, ByVal v2 As System.Version) As Boolean
'Public Function op_Equality(ByVal v1 As System.Version, ByVal v2 As System.Version) As cBoolean
'  Set op_Equality = New cBoolean
'  If v1.CompareTo(v2) = 0 Then op_Equality = True
'End Function

'Public Shared Function op_GreaterThan(ByVal v1 As System.Version, ByVal v2 As System.Version) As Boolean
'Public Function op_GreaterThan(ByVal v1 As System.Version, ByVal v2 As System.Version) As cBoolean
'  Set op_GreaterThan = New cBoolean
'  If v1.CompareTo(v2) > 0 Then op_GreaterThan = True
'End Function
Public Function GreaterThan(other As Version) As Boolean
    GreaterThan = other.Major < m_Major: If GreaterThan Then Exit Function 'return -1
    GreaterThan = other.Major < m_Major: If GreaterThan Then Exit Function 'return -1
End Function

'Public Shared Function op_GreaterThanOrEqual(ByVal v1 As System.Version, ByVal v2 As System.Version) As Boolean
'Public Static Function op_GreaterThanOrEqual(ByVal v1 As System.Version, ByVal v2 As System.Version) As cBoolean
'  Set op_GreaterThanOrEqual = New cBoolean
'  If v1.CompareTo(v2) >= 0 Then op_GreaterThanOrEqual = True
'End Function

'Public Shared Function op_Inequality(ByVal v1 As System.Version, ByVal v2 As System.Version) As Boolean
'Public Static Function op_Inequality(ByVal v1 As System.Version, ByVal v2 As System.Version) As cBoolean
'  Set op_Inequality = New cBoolean
'  If v1.CompareTo(v2) <> 0 Then op_Inequality = True
'End Function

'Public Shared Function op_LessThan(ByVal v1 As System.Version, ByVal v2 As System.Version) As Boolean
'Public Static Function op_LessThan(ByVal v1 As System.Version, ByVal v2 As System.Version) As cBoolean
'  Set op_LessThan = New cBoolean
'  If v1.CompareTo(v2) < 0 Then op_LessThan = True
'End Function

'Public Shared Function op_LessThanOrEqual(ByVal v1 As System.Version, ByVal v2 As System.Version) As Boolean
'Public Static Function op_LessThanOrEqual(ByVal v1 As System.Version, ByVal v2 As System.Version) As cBoolean
'  Set op_LessThanOrEqual = New cBoolean
'  If v1.CompareTo(v2) <= 0 Then op_LessThanOrEqual = True
'End Function

'Public Overrides Function ToString() As String
'Public Function ToString(ByVal fieldCount As Integer) As String
Public Function ToStr(Optional ByVal fieldCount As Long) As String
    'show minor.major always
    Dim s As String: s = CStr(m_Major) & "." & CStr(m_Minor)
    'show Build and Revision only if they are greater or equal 0
    If m_Build > -1 Then s = s & "." & CStr(m_Build)
    If m_Revis > -1 Then s = s & "." & CStr(m_Revis)
    ToStr = s
End Function

Public Property Get Major() As Long
    Major = m_Major
End Property

Public Property Get Minor() As Long
    Minor = m_Minor
End Property

Public Property Get Build() As Long
    Build = m_Build
End Property

Public Property Get Revision() As Long
    Revision = m_Revis
End Property


