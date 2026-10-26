VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TOrderedPolygon3D"
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
'this type is the same as TCHullPoint3D
Private Type TOrdPolyPoint3D
  X   As Double 'Double;
  Y   As Double 'Double;
  z   As Double 'Double;
  ang As Double 'Double;
End Type
Private Point() As TOrdPolyPoint3D

Private Sub Class_Initialize()
  '
End Sub
Friend Sub OrderedPolygon(Pnt() As TPoint3D)
Dim i As Long, LargestX As Double, LargestZ As Double
  ReDim Point(UBound(Pnt))
  LargestX = 0
  LargestZ = 0
  For i = 0 To UBound(Pnt)
    Point(i).X = Pnt(i).X
    Point(i).Y = Pnt(i).Y
    Point(i).z = Pnt(i).z
    Point(i).ang = 0#
    If Point(i).Y < Point(0).Y Then Call Swap(0, i)
    If Point(i).X > LargestX Then LargestX = Point(i).X
    If Point(i).z > LargestZ Then LargestZ = Point(i).z
  Next
  For i = 1 To UBound(Pnt)
    Point(i).ang = MFastGeo.VertexAngle3DXY(LargestX, Point(0).Y, LargestZ, Point(0).X, Point(0).Y, Point(0).z, Point(i).X, Point(i).Y, Point(i).z)
  Next
  Call RQSort(1, UBound(Point))
  For i = 0 To UBound(Pnt)
    Pnt(i).X = Point(i).X
    Pnt(i).Y = Point(i).Y
    Pnt(i).z = Point(i).z
  Next
End Sub
Private Sub Swap(i As Long, j As Long)
Dim Temp As TOrdPolyPoint3D
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
  '  apply quick-sort to both left
  '  and right sides of pivoting
  '*)
  Call RQSort(Left, i - 1)
  Call RQSort(i + 1, Right)
End Sub
Private Function Partition(Left As Long, Right As Long) As Long
Dim i As Long, j As Long, Middle As Long, Pivot As TOrdPolyPoint3D
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
    Do While (Pivot.ang <= Point(j).ang)
      Call Dec(j)
      If j - 1 < 0 Then Exit Function
    Loop
    If i < j Then Call Swap(i, j)
  Loop Until i >= j

  Call Swap(i, Right)
  Partition = i
End Function

