VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TOrderedPolygon2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Double Precision!
Option Base 0
#If defFastGeoAXdll Then
Public FastGeo As FastGeo
#End If
'this type is the same as TCHullPoint2D
Private Type TOrdPolyPoint2D
  X   As Double 'Double;
  Y   As Double 'Double;
  ang As Double 'Double;
End Type
Private Point() As TOrdPolyPoint2D
Private Sub Class_Initialize()
  '
End Sub
Friend Sub OrderedPolygon(Pnt() As TPoint2D)
Dim i As Long, j As Long
  ReDim Point(UBound(Pnt))
  j = 0
  For i = 0 To UBound(Point)
    Point(i).X = Pnt(i).X
    Point(i).Y = Pnt(i).Y
    Point(i).ang = 0#
    If Point(i).Y < Point(j).Y Then
      j = i
    ElseIf Point(i).Y = Point(j).Y Then
      If Point(i).X < Point(j).X Then
        j = i
      End If
    End If
  Next
  Call Swap(0, j)
  For i = 1 To UBound(Point)
    Point(i).ang = MFastGeo.CartesianAngle(Point(i).X - Point(0).X, Point(i).Y - Point(0).Y)
  Next
  Call RQSort(1, UBound(Point))
  For i = 0 To UBound(Point)
    Pnt(i).X = Point(i).X
    Pnt(i).Y = Point(i).Y
  Next
End Sub
Private Sub Swap(i As Long, j As Long)
Dim Temp As TOrdPolyPoint2D
  Temp = Point(i)
  Point(i) = Point(j)
  Point(j) = Temp
End Sub
Private Sub RQSort(Left As Long, Right As Long)
Dim i As Long
  '(* ends have crossed over pivot *)
  If Right <= Left Then Exit Sub
  '(* paritition array *)
  i = Partition(Left, Right)
  '(*
  '   apply quick-sort to both left
  '   and right sides of pivoting
  '*)
  Call RQSort(Left, i - 1)
  Call RQSort(i + 1, Right)
End Sub
Private Function Partition(Left As Long, Right As Long) As Long
Dim i As Long, j As Long, Middle As Long, Pivot As TOrdPolyPoint2D
  i = Left
  j = Right
  Middle = (Left + Right) \ 2
  '{ Median of 3 Pivot Selection }
  If Point(Middle).ang < Point(Left).ang Then Call Swap(Left, Middle)
  If Point(Right).ang < Point(Middle).ang Then Call Swap(Right, Middle)
  If Point(Middle).ang < Point(Left).ang Then Call Swap(Left, Middle)
  Pivot.ang = Point(Right).ang
  Do
    Do While (Pivot.ang >= Point(i).ang) And (i < Right)
      Call Inc(i)
    Loop
    Do While Pivot.ang <= Point(j).ang
      Call Dec(j)
      If j - 1 < 0 Then Exit Do
    Loop
    If i < j Then Call Swap(i, j)
  Loop Until i >= j
  Call Swap(i, Right)
  Partition = i
End Function
