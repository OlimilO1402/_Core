Attribute VB_Name = "MFastGeo"
Option Explicit
'Double Precision!
'(*************************************************************************)
'(*                                                                       *)
'(*                             FASTGEO                                   *)
'(*                                                                       *)
'(*                2D/3D Computational Geometry Algorithms                *)
'(*                        Release Version 5.0.1                          *)
'(*                                                                       *)
'(* Author: Arash Partow 1997-2006                                        *)
'(* URL: http://fastgeo.partow.net                                        *)
'(*      http://www.partow.net/projects/fastgeo/index.html                *)
'(*                                                                       *)
'(* Copyright notice:                                                     *)
'(* Free use of the FastGEO computational geometry library is permitted   *)
'(* under the guidelines and in accordance with the most current version  *)
'(* of the Common Public License.                                         *)
'(* http://www.opensource.org/licenses/cpl.php                            *)
'(*                                                                       *)
'(*************************************************************************)
'
'Originally written in Delphi
'brouhgt to the VB-World by MBO-Ing.com
Public Const VersionInformation As String = "FastGEO Version 5.0.1"
Public Const AuthorInformation As String = "Arash Partow (1997-2006)"
Public Const EpochInformation As String = "Lambda-Phi"
Public Const RecentUpdate As String = "10-02-2006"
Public Const LCID As String = "$10-02-2006:FEDEDB4632780C2FAE$"
Public Const Precision As String = "Double Precision"
Public SystemEpsilon As Double
'(* 2D/3D Portal Definition *)
Public MaximumX As Double
Public MinimumX As Double
Public MaximumY As Double
Public MinimumY As Double
Public MaximumZ As Double
Public MinimumZ As Double

Private SinTable() As Double
Private CosTable() As Double
Private TanTable() As Double

'###########################################################################
'Orientation2D
Public Function Orientation2D(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double) As Long
Dim Orin As Double
  '(* Determinant of the 3 points *)
  Orin = (X2 - X1) * (Py - Y1) - (Px - X1) * (Y2 - Y1)
  If Orin > 0# Then
    Orientation2D = LeftHandSide          '(* Orientaion is to the left-hand side  *)
  ElseIf Orin < 0# Then
    Orientation2D = RightHandSide         '(* Orientaion is to the right-hand side *)
  Else
    Orientation2D = CollinearOrientation  '(* Orientaion is neutral aka collinear  *)
  End If
End Function
Public Function Orientation3D(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, Px As Double, Py As Double, Pz As Double) As Long
Dim Px1 As Double, Px2 As Double, Px3 As Double, Py1 As Double, Py2 As Double, Py3 As Double, Pz1 As Double, Pz2 As Double, Pz3 As Double, Orin As Double
  Px1 = X1 - Px
  Px2 = X2 - Px
  Px3 = X3 - Px
  
  Py1 = Y1 - Py
  Py2 = Y2 - Py
  Py3 = Y3 - Py
  
  Pz1 = z1 - Pz
  Pz2 = z2 - Pz
  Pz3 = z3 - Pz
  
  Orin = Px1 * (Py2 * Pz3 - Pz2 * Py3) + _
         Px2 * (Py3 * Pz1 - Pz3 * Py1) + _
         Px3 * (Py1 * Pz2 - Pz1 * Py2)
           
  If Orin < 0# Then
    Orientation3D = -1           '(* Orientaion is below plane                      *)
  ElseIf Orin > 0# Then
    Orientation3D = 1            '(* Orientaion is above plane                      *)
  Else
    Orientation3D = 0            '(* Orientaion is coplanar to plane if Result is 0 *)
  End If
End Function
'(* End of Orientation *)

'RobustOrientation
Public Function RobustOrientation2D(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double) As Long
Dim Orin As Double
  '(* Linear determinant of the 3 points *)
  Orin = (X2 - X1) * (Py - Y1) - (Px - X1) * (Y2 - Y1)
  '(*
  '  Calculation Policy:
  '  if |Orin - Orin`| < Epsilon then Orin` is assumed to be equal to zero.
  'Where:
  '   Orin : is the "real" mathematically precise orientation value, using infinite
  '          precision arithmetic(hypothetical)
  '   Orin`: is the calculated imprecise orientation value, using finite precision arithmetic
  '*)
  If IsEqual(Orin, 0#) Then
    RobustOrientation2D = 0                '(* Orientaion is neutral aka collinear  *)
  ElseIf Orin < 0# Then
    RobustOrientation2D = RightHandSide    '(* Orientaion is to the right-hand side *)
  Else
    RobustOrientation2D = LeftHandSide     '(* Orientaion is to the left-hand side  *)
  End If
End Function
Public Function RobustOrientation3D(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, Px As Double, Py As Double, Pz As Double) As Long
Dim Px1 As Double, Px2 As Double, Px3 As Double, Py1 As Double, Py2 As Double, Py3 As Double, Pz1 As Double, Pz2 As Double, Pz3 As Double, Orin As Double
  Px1 = X1 - Px
  Px2 = X2 - Px
  Px3 = X3 - Px

  Py1 = Y1 - Py
  Py2 = Y2 - Py
  Py3 = Y3 - Py
  
  Pz1 = z1 - Pz
  Pz2 = z2 - Pz
  Pz3 = z3 - Pz

  Orin = Px1 * (Py2 * Pz3 - Pz2 * Py3) + _
         Px2 * (Py3 * Pz1 - Pz3 * Py1) + _
         Px3 * (Py1 * Pz2 - Pz1 * Py2)

  If IsEqual(Orin, 0#) Then
    RobustOrientation3D = 0              '(* Orientaion is coplanar to plane if Result is 0 *)
  ElseIf Orin < 0# Then
    RobustOrientation3D = -1             '(* Orientaion is below plane                      *)
  Else
    RobustOrientation3D = 1              '(* Orientaion is above plane                      *)
  End If
End Function
'(* End of Robust Orientation *)

Public Function Orientation2DPXY(Point1 As TPoint2D, Point2 As TPoint2D, Px As Double, Py As Double) As Long
  Orientation2DPXY = Orientation2D(Point1.X, Point1.Y, Point2.X, Point2.Y, Px, Py)
End Function
Public Function Orientation2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Long
  Orientation2DP = Orientation2D(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function Orientation2DLP(Line As TLine2D, Point As TPoint2D) As Long
  Orientation2DLP = Orientation2D(Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y, Point.X, Point.Y)
End Function
Public Function Orientation2DSP(Segment As TSegment2D, Point As TPoint2D) As Long
  Orientation2DSP = Orientation2D(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Point.X, Point.Y)
End Function
Public Function Orientation3DPXY(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Px As Double, Py As Double, Pz As Double) As Long
  Orientation3DPXY = Orientation3D(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Px, Py, Pz)
End Function
Public Function Orientation3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Point4 As TPoint3D) As Long
  Orientation3DP = Orientation3D(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Point4.X, Point4.Y, Point4.z)
End Function
Public Function Orientation3DTP(Triangle As TTriangle3D, Point As TPoint3D) As Long
  Orientation3DTP = Orientation3DP(Triangle.P(0), Triangle.P(1), Triangle.P(2), Point)
End Function
'(* End of Orientation *)

'Signed
Public Function Signed2D(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double) As Double
  Signed2D = (X2 - X1) * (Py - Y1) - (Px - X1) * (Y2 - Y1)
End Function
Public Function Signed3D(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, Px As Double, Py As Double, Pz As Double) As Double
Dim Px1 As Double, Px2 As Double, Px3 As Double, Py1 As Double, Py2 As Double, Py3 As Double, Pz1 As Double, Pz2 As Double, Pz3 As Double
  Px1 = X1 - Px:  Px2 = X2 - Px:  Px3 = X3 - Px
  Py1 = Y1 - Py:  Py2 = Y2 - Py:  Py3 = Y3 - Py
  Pz1 = z1 - Pz:  Pz2 = z2 - Pz:  Pz3 = z3 - Pz
  Signed3D = Px1 * (Py2 * Pz3 - Pz2 * Py3) + Px2 * (Py3 * Pz1 - Pz3 * Py1) + Px3 * (Py1 * Pz2 - Pz1 * Py2)
End Function
Public Function Signed2DPxy(Point1 As TPoint2D, Point2 As TPoint2D, Px As Double, Py As Double) As Double
  Signed2DPxy = Signed2D(Point1.X, Point1.Y, Point2.X, Point2.Y, Px, Py)
End Function
Public Function Signed2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Double
  Signed2DP = Signed2D(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function Signed2DLP(Line As TLine2D, Point As TPoint2D) As Double
  Signed2DLP = Signed2D(Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y, Point.X, Point.Y)
End Function
Public Function Signed2DSP(Segment As TSegment2D, Point As TPoint2D) As Double
  Signed2DSP = Signed2D(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Point.X, Point.Y)
End Function
Public Function Signed3DPxyz(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Px As Double, Py As Double, Pz As Double) As Double
  Signed3DPxyz = Signed3D(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Px, Py, Pz)
End Function
Public Function Signed3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Point4 As TPoint3D) As Double
  Signed3DP = Signed3D(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Point4.X, Point4.Y, Point4.z)
End Function
Public Function Signed3DTP(Triangle As TTriangle3D, Point As TPoint3D) As Double
  Signed3DTP = Signed3DP(Triangle.P(0), Triangle.P(1), Triangle.P(2), Point)
End Function
'(* End of Signed *)

'Collinear
Public Function Collinear2D(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Boolean
  Collinear2D = IsEqual((X2 - X1) * (Y3 - Y1), (X3 - X1) * (Y2 - Y1))
End Function
Public Function Collinear2DEps(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Epsilon As Double) As Boolean
  Collinear2DEps = IsEqualEps((X2 - X1) * (Y3 - Y1), (X3 - X1) * (Y2 - Y1), Epsilon)
End Function
Public Function Collinear3D(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As Boolean
Dim Dx1 As Double, Dx2 As Double, Dy1 As Double, Dy2 As Double, Dz1 As Double, Dz2 As Double, cX As Double, cY As Double, cz As Double
  '(* Find the difference between the 2 points P2 and P3 to P1 *)
  Dx1 = X2 - X1
  Dy1 = Y2 - Y1
  Dz1 = z2 - z1

  Dx2 = X3 - X1
  Dy2 = Y3 - Y1
  Dz2 = z3 - z1

  '(* Perform a 3d cross product *)
  cX = (Dy1 * Dz2) - (Dy2 * Dz1)
  cY = (Dx2 * Dz1) - (Dx1 * Dz2)
  cz = (Dx1 * Dy2) - (Dx2 * Dy1)
  Collinear3D = IsEqual(cX * cX + cY * cY + cz * cz, 0#)
End Function
Public Function Collinear2DP(PointA As TPoint2D, PointB As TPoint2D, PointC As TPoint2D) As Boolean
  Collinear2DP = Collinear2D(PointA.X, PointA.Y, PointB.X, PointB.Y, PointC.X, PointC.Y)
End Function
Public Function Collinear3DP(PointA As TPoint3D, PointB As TPoint3D, PointC As TPoint3D) As Boolean
  Collinear3DP = Collinear3D(PointA.X, PointA.Y, PointA.z, PointB.X, PointB.Y, PointB.z, PointC.X, PointC.Y, PointC.z)
End Function
'(* End of Collinear *)

'RobustCollinear
Public Function RobustCollinear(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Optional Epsilon As Double = Epsilon_High) As Boolean
Dim Leydist1 As Double, Leydist2 As Double, Leydist3 As Double
  Leydist1 = LayDistance2DXY(X1, Y1, X2, Y2)
  Leydist2 = LayDistance2DXY(X2, Y2, X3, Y3)
  Leydist3 = LayDistance2DXY(X3, Y3, X1, Y1)
  If Leydist1 >= Leydist2 Then
    If Leydist1 >= Leydist3 Then
      RobustCollinear = IsEqualEps(MinimumDistanceFromPointToLine2DXY(X3, Y3, X1, Y1, X2, Y2), 0#, Epsilon)
    Else
      RobustCollinear = IsEqualEps(MinimumDistanceFromPointToLine2DXY(X2, Y2, X3, Y3, X1, Y1), 0#, Epsilon)
    End If
  ElseIf Leydist2 >= Leydist3 Then
    RobustCollinear = IsEqualEps(MinimumDistanceFromPointToLine2DXY(X1, Y1, X2, Y2, X3, Y3), 0#, Epsilon)
  Else
    RobustCollinear = IsEqualEps(MinimumDistanceFromPointToLine2DXY(X2, Y2, X3, Y3, X1, Y1), 0#, Epsilon)
  End If
End Function
Public Function RobustCollinearP(PointA As TPoint2D, PointB As TPoint2D, PointC As TPoint2D, Optional Epsilon As Double = Epsilon_High) As Boolean
  RobustCollinearP = RobustCollinear(PointA.X, PointA.Y, PointB.X, PointB.Y, PointC.X, PointC.Y, Epsilon)
End Function
'(* End of Robust Collinear *)

'Coplanar
Public Function Coplanar3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double) As Boolean
  Coplanar3DXY = (Orientation3D(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3, X4, Y4, z4) = CoplanarOrientation)
End Function
Public Function Coplanar3DP(PointA As TPoint3D, PointB As TPoint3D, PointC As TPoint3D, PointD As TPoint3D) As Boolean
  Coplanar3DP = (Orientation3DP(PointA, PointB, PointC, PointD) = CoplanarOrientation)
End Function
'(* End of Coplanar *)

'IsPointCollinear
Public Function IsPointCollinear2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double, Optional Robust As Boolean = False) As Boolean
  '(*
  '  This method will return true If the point (px,py) is collinear
  '  to points (x1,y1) and (x2,y2) and exists on the segment A(x1,y1)->B(x2,y2)
  '*)
  If (((X1 <= Px) And (Px <= X2)) Or ((X2 <= Px) And (Px <= X1))) And _
     (((Y1 <= Py) And (Py <= Y2)) Or ((Y2 <= Py) And (Py <= Y1))) Then
    If Robust Then
      IsPointCollinear2DXY = RobustCollinear(X1, Y1, X2, Y2, Px, Py)
    Else
      IsPointCollinear2DXY = Collinear2D(X1, Y1, X2, Y2, Px, Py)
    End If
  Else
    IsPointCollinear2DXY = False
  End If
End Function
Public Function IsPointCollinear2DP(PointA As TPoint2D, PointB As TPoint2D, PointC As TPoint2D, Optional Robust As Boolean = False) As Boolean
 '(*
 '  This method will return true Iff the pointC is collinear
 '  to points A and B and exists on the segment A->B or B->A
 '*)
  IsPointCollinear2DP = IsPointCollinear2DXY(PointA.X, PointA.Y, PointB.X, PointB.Y, PointC.X, PointC.Y, Robust)
End Function
Public Function IsPointCollinear2DPXY(PointA As TPoint2D, PointB As TPoint2D, Px As Double, Py As Double, Optional Robust As Boolean = False) As Boolean
  IsPointCollinear2DPXY = IsPointCollinear2DXY(PointA.X, PointA.Y, PointB.X, PointB.Y, Px, Py, Robust)
End Function
Public Function IsPointCollinear2DSP(Segment As TSegment2D, PointC As TPoint2D, Optional Robust As Boolean = False) As Boolean
  IsPointCollinear2DSP = IsPointCollinear2DP(Segment.P(0), Segment.P(1), PointC, Robust)
End Function
Public Function IsPointCollinear3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, Px As Double, Py As Double, Pz As Double) As Boolean
 '(*
 '  This method will return true Iff the point (px,py,pz) is collinear
 '  to points (x1,y1,z1) and (x2,y2,z2) and exists on the segment A(x1,y1,z1)->B(x2,y2,z2)
 '*)
 If (((X1 <= Px) And (Px <= X2)) Or ((X2 <= Px) And (Px <= X1))) And _
    (((Y1 <= Py) And (Py <= Y2)) Or ((Y2 <= Py) And (Py <= Y1))) And _
    (((z1 <= Pz) And (Pz <= z2)) Or ((z2 <= Pz) And (Pz <= z1))) Then
   IsPointCollinear3DXY = Collinear3D(X1, Y1, z1, X2, Y2, z2, Px, Py, Pz)
 Else
   IsPointCollinear3DXY = False
 End If
End Function
Public Function IsPointCollinear3DP(PointA As TPoint3D, PointB As TPoint3D, PointC As TPoint3D) As Boolean
  '(*
  '  This method will return true Iff the pointC is collinear
  '  to points A and B and exists on the segment A->B or B->A
  '*)
  IsPointCollinear3DP = IsPointCollinear3DXY(PointA.X, PointA.Y, PointA.z, PointB.X, PointB.Y, PointB.z, PointC.X, PointC.Y, PointC.z)
End Function
Public Function IsPointCollinear3DSP(Segment As TSegment3D, PointC As TPoint3D) As Boolean
  IsPointCollinear3DSP = IsPointCollinear3DP(Segment.P(0), Segment.P(1), PointC)
End Function
'(* End of IsPointCollinear *)

'IsPointOnRightSide
Public Function IsPointOnRightSide2DXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Boolean
  IsPointOnRightSide2DXY = (((X2 - X1) * (Py - Y1)) < ((Px - X1) * (Y2 - Y1)))
End Function
Public Function IsPointOnRightSide2DXYS(X As Double, Y As Double, Segment As TSegment2D) As Boolean
  IsPointOnRightSide2DXYS = IsPointOnRightSide2DXY(X, Y, Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y)
End Function
Public Function IsPointOnRightSide2DPS(Point As TPoint2D, Segment As TSegment2D) As Boolean
  IsPointOnRightSide2DPS = IsPointOnRightSide2DXYS(Point.X, Point.Y, Segment)
End Function
Public Function IsPointOnRightSide2DXYL(X As Double, Y As Double, Line As TLine2D) As Boolean
  IsPointOnRightSide2DXYL = IsPointOnRightSide2DXY(X, Y, Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y)
End Function
Public Function IsPointOnRightSide2DPL(Point As TPoint2D, Line As TLine2D) As Boolean
  IsPointOnRightSide2DPL = IsPointOnRightSide2DXYL(Point.X, Point.Y, Line)
End Function
'(* End of IsPointOnRightSide *)

'IsPointOnLeftSide
Public Function IsPointOnLeftSide2DXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Boolean
  IsPointOnLeftSide2DXY = ((X2 - X1) * (Py - Y1) > (Px - X1) * (Y2 - Y1))
End Function
Public Function IsPointOnLeftSide2DXYS(X As Double, Y As Double, Segment As TSegment2D) As Boolean
  IsPointOnLeftSide2DXYS = IsPointOnLeftSide2DXY(X, Y, Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y)
End Function
Public Function IsPointOnLeftSide2DPS(Point As TPoint2D, Segment As TSegment2D) As Boolean
  IsPointOnLeftSide2DPS = IsPointOnLeftSide2DXYS(Point.X, Point.Y, Segment)
End Function
Public Function IsPointOnLeftSide2DXYL(X As Double, Y As Double, Line As TLine2D) As Boolean
  IsPointOnLeftSide2DXYL = IsPointOnLeftSide2DXY(X, Y, Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y)
End Function
Public Function IsPointOnLeftSide2DPL(Point As TPoint2D, Line As TLine2D) As Boolean
  IsPointOnLeftSide2DPL = IsPointOnLeftSide2DXYL(Point.X, Point.Y, Line)
End Function
'(* End of IsPointOnLeftSide *)

Public Function RotationCCW(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double) As Long
  If ((Px - X1) * (Y2 - Y1)) > ((X2 - X1) * (Py - Y1)) Then
    RotationCCW = 1
  Else
    RotationCCW = -1
  End If
End Function
'(* End of RotationCCW *)
Public Function RotationCW(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double) As Long
  If ((X2 - X1) * (Py - Y1)) > ((Px - X1) * (Y2 - Y1)) Then
    RotationCW = 1
  Else
    RotationCW = -1
  End If
End Function
'(* End of RotationCW *)

'Intersect
Public Function Intersect2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Boolean
Dim UpperX  As Double, UpperY  As Double, LowerX   As Double, LowerY   As Double
Dim ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double
Dim D As Double, F As Double, E As Double
  Intersect2DXY = False

  ax = X2 - X1
  bx = X3 - X4

  If ax < 0# Then
    LowerX = X2
    UpperX = X1
  Else
    UpperX = X2
    LowerX = X1
  End If

  If bx > 0# Then
    If (UpperX < X4) Or (X3 < LowerX) Then
      Exit Function
    End If
  ElseIf (UpperX < X3) Or (X4 < LowerX) Then
    Exit Function
  End If
  
  ay = Y2 - Y1
  by = Y3 - Y4

  If ay < 0# Then
    LowerY = Y2
    UpperY = Y1
  Else
    UpperY = Y2
    LowerY = Y1
  End If

  If by > 0# Then
    If (UpperY < Y4) Or (Y3 < LowerY) Then
      Exit Function
    End If
  ElseIf (UpperY < Y3) Or (Y4 < LowerY) Then
    Exit Function
  End If

  cX = X1 - X3
  cY = Y1 - Y3
  D = (by * cX) - (bx * cY)
  F = (ay * bx) - (ax * by)

  If F > 0# Then
    If (D < 0#) Or (D > F) Then
      Exit Function
    End If
  ElseIf (D > 0#) Or (D < F) Then
    Exit Function
  End If

  E = (ax * cY) - (ay * cX)

  If F > 0# Then
    If (E < 0#) Or (E > F) Then
      Exit Function
    End If
  ElseIf (E > 0#) Or (E < F) Then
    Exit Function
  End If
  
  Intersect2DXY = True
  
  '(*
  '
  'Simple method, yet not so accurate for certain situations and a little more
  'inefficient (roughly 19.5%).
  'Result = (
  '           ((Orientation(x1,y1, x2,y2, x3,y3) * Orientation(x1,y1, x2,y2, x4,y4)) <= 0) and
  '           ((Orientation(x3,y3, x4,y4, x1,y1) * Orientation(x3,y3, x4,y4, x2,y2)) <= 0)
  '          )
  '*)
End Function
Public Function Intersect2DXYi(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, ix As Double, iy As Double) As Boolean
Dim UpperX As Double, UpperY As Double, LowerX As Double, LowerY As Double
Dim ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double
Dim D As Double, F As Double, E As Double, Ratio As Double
  
  Intersect2DXYi = False
  
  ax = X2 - X1
  bx = X3 - X4
  
  If ax < 0# Then
    LowerX = X2
    UpperX = X1
  Else
    UpperX = X2
    LowerX = X1
  End If
  
  If bx > 0# Then
    If (UpperX < X4) Or (X3 < LowerX) Then
      Exit Function
    End If
  ElseIf (UpperX < X3) Or (X4 < LowerX) Then
    Exit Function
  End If
  
  ay = Y2 - Y1
  by = Y3 - Y4
  
  If ay < 0# Then
    LowerY = Y2
    UpperY = Y1
  Else
    UpperY = Y2
    LowerY = Y1
  End If
  
  If by > 0# Then
    If (UpperY < Y4) Or (Y3 < LowerY) Then
      Exit Function
    End If
  ElseIf (UpperY < Y3) Or (Y4 < LowerY) Then
    Exit Function
  End If
  
  cX = X1 - X3
  cY = Y1 - Y3
  D = (by * cX) - (bx * cY)
  F = (ay * bx) - (ax * by)
  
  If F > 0# Then
    If (D < 0#) Or (D > F) Then
      Exit Function
    End If
  ElseIf (D > 0#) Or (D < F) Then
    Exit Function
  End If
  
  E = (ax * cY) - (ay * cX)
  
  If F > 0# Then
    If (E < 0#) Or (E > F) Then
      Exit Function
    End If
  ElseIf (E > 0#) Or (E < F) Then
    Exit Function
  End If
  Intersect2DXYi = True
  
  '(*
  '
  '  From IntersectionPoint Routine
  '
  '  dx1 = x2 - x1 ->  Ax
  '  dx2 = x4 - x3 -> -Bx
  '  dx3 = x1 - x3 ->  Cx
  '
  '  dy1 = y2 - y1 ->  Ay
  '  dy2 = y1 - y3 ->  Cy
  '  dy3 = y4 - y3 -> -By
  '
  '*)
  
  Ratio = (ax * -by) - (ay * -bx)
  
  If NotEqual(Ratio, 0#) Then
    Ratio = ((cY * -bx) - (cX * -by)) / Ratio
    ix = X1 + (Ratio * ax)
    iy = Y1 + (Ratio * ay)
  Else
    '//if Collinear(x1,y1,x2,y2,x3,y3) then
    If IsEqual((ax * -cY), (-cX * ay)) Then
      ix = X3
      iy = Y3
    Else
      ix = X4
      iy = Y4
    End If
  End If
End Function

Public Function Intersect2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D) As Boolean
  Intersect2DP = Intersect2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Point4.X, Point4.Y)
End Function
Public Function Intersect2DS(Segment1 As TSegment2D, Segment2 As TSegment2D) As Boolean
  Intersect2DS = Intersect2DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1))
End Function
Public Function Intersect2DSXY(Segment1 As TSegment2D, Segment2 As TSegment2D, ix As Double, iy As Double) As Boolean
  Intersect2DSXY = Intersect2DXYi(Segment1.P(0).X, Segment1.P(0).Y, Segment1.P(1).X, Segment1.P(1).Y, Segment2.P(0).X, Segment2.P(0).Y, Segment2.P(1).X, Segment2.P(1).Y, ix, iy)
End Function
Public Function Intersect2DSR(Segment As TSegment2D, aRectangle As TRectangle) As Boolean
  Intersect2DSR = RectangleToRectangleIntersectR(aRectangle, EquateRectangleP(Segment.P(0), Segment.P(1)))
End Function
Public Function Intersect2DST(Segment As TSegment2D, Triangle As TTriangle2D) As Boolean
  Intersect2DST = Intersect2DS(Segment, TriangleEdge2D(Triangle, 1)) Or _
                  Intersect2DS(Segment, TriangleEdge2D(Triangle, 2)) Or _
                  Intersect2DS(Segment, TriangleEdge2D(Triangle, 3)) Or _
               PointInTriangle(Segment.P(0), Triangle) Or _
               PointInTriangle(Segment.P(1), Triangle)
End Function
Public Function Intersect2DSQ(Segment As TSegment2D, Quadix As TQuadix2D) As Boolean
  Intersect2DSQ = Intersect2DST(Segment, EquateTriangle2DP(Quadix.P(0), Quadix.P(1), Quadix.P(2))) Or _
                  Intersect2DST(Segment, EquateTriangle2DP(Quadix.P(0), Quadix.P(2), Quadix.P(3)))
End Function
Public Function Intersect2DSL(Segment As TSegment2D, Line As TLine2D) As Boolean
  Intersect2DSL = (Orientation2DLP(Line, Segment.P(0)) * Orientation2DLP(Line, Segment.P(1)) <= 0)
End Function
Public Function Intersect2DSCi(Segment As TSegment2D, aCircle As TCircle) As Boolean
Dim Px As Double, Py As Double
  Call ClosestPointOnSegmentFromPointS2DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, aCircle.X, aCircle.Y, Px, Py)
  Intersect2DSCi = (LayDistance2DXY(Px, Py, aCircle.X, aCircle.Y) <= (aCircle.Radius * aCircle.Radius))
End Function
Public Function Intersect2DLT(Line As TLine2D, Triangle As TTriangle2D) As Boolean
Dim Or1 As Long, Or2 As Long
  Intersect2DLT = True
  Or1 = Orientation2DP(Line.P(0), Line.P(1), Triangle.P(0))
  If Or1 = 0 Then Exit Function
  Or2 = Orientation2DP(Line.P(0), Line.P(1), Triangle.P(1))
  If (Or2 <> Or1) Then Exit Function
  Or2 = Orientation2DP(Line.P(0), Line.P(1), Triangle.P(2))
  Intersect2DLT = Or2 <> Or1
End Function
Public Function Intersect2DLQ(Line As TLine2D, Quadix As TQuadix2D) As Boolean
Dim Or1 As Long, Or2 As Long
  Intersect2DLQ = True
  Or1 = Orientation2DP(Line.P(0), Line.P(1), Quadix.P(0))
  If Or1 = 0 Then Exit Function
  Or2 = Orientation2DP(Line.P(0), Line.P(1), Quadix.P(1))
  If (Or2 <> Or1) Then Exit Function
  Or2 = Orientation2DP(Line.P(0), Line.P(1), Quadix.P(2))
  If (Or2 <> Or1) Then Exit Function
  Or2 = Orientation2DP(Line.P(0), Line.P(1), Quadix.P(3))
  Intersect2DLQ = Or2 <> Or1
End Function
Public Function Intersect2DLCi(Line As TLine2D, aCircle As TCircle) As Boolean
Dim X1  As Double, Y1  As Double, X2   As Double, Y2   As Double
'  (*
'    It is assumed that an intersection of a circle by a line
'    is either a full intersection (2 points), partial intersection
'    (1 point), or tangential.
'    Anything else will Result in a false output.
'  *)
  X1 = Line.P(0).X - aCircle.X
  Y1 = Line.P(0).Y - aCircle.Y
  X2 = Line.P(1).X - aCircle.X
  Y2 = Line.P(1).Y - aCircle.Y
  Intersect2DLCi = GreaterThanOrEqual(((aCircle.Radius * aCircle.Radius) * LayDistance2DXY(X1, Y1, X2, Y2) - Sqr(X1 * Y2 - X2 * Y1)), 0#)
End Function
Public Function Intersect2DTCi(Triangle As TTriangle2D, aCircle As TCircle) As Boolean
  Intersect2DTCi = PointInCircle(ClosestPointOnTriangleFromPoint2DTXY(Triangle, aCircle.X, aCircle.Y), aCircle)
End Function
Public Function Intersect2DTR(Triangle As TTriangle2D, aRectangle As TRectangle) As Boolean
  Intersect2DTR = Intersect2DSR(EquateSegment2DP(Triangle.P(0), Triangle.P(1)), aRectangle) Or _
                  Intersect2DSR(EquateSegment2DP(Triangle.P(1), Triangle.P(2)), aRectangle) Or _
                  Intersect2DSR(EquateSegment2DP(Triangle.P(2), Triangle.P(0)), aRectangle)
End Function
Public Function Intersect2DR(Rectangle1 As TRectangle, Rectangle2 As TRectangle) As Boolean
  Intersect2DR = RectangleToRectangleIntersectR(Rectangle1, Rectangle2)
End Function
Public Function Intersect2DTT(Triangle1 As TTriangle2D, Triangle2 As TTriangle2D) As Boolean
  Intersect2DTT = False
  If IsEqual(MinimumDistanceFromPointToTriangle(Triangle1.P(0), Triangle2), 0#) Or _
     IsEqual(MinimumDistanceFromPointToTriangle(Triangle2.P(0), Triangle1), 0#) Then
    Intersect2DTT = True
    Exit Function
  End If
  If IsEqual(MinimumDistanceFromPointToTriangle(Triangle1.P(1), Triangle2), 0#) Or _
     IsEqual(MinimumDistanceFromPointToTriangle(Triangle2.P(1), Triangle1), 0#) Then
    Intersect2DTT = True
    Exit Function
  End If
  If IsEqual(MinimumDistanceFromPointToTriangle(Triangle1.P(2), Triangle2), 0#) Or _
     IsEqual(MinimumDistanceFromPointToTriangle(Triangle2.P(2), Triangle1), 0#) Then
    Intersect2DTT = True
    Exit Function
  End If
End Function
Public Function Intersect2DRCi(Rectangle As TRectangle, aCircle As TCircle) As Boolean
  Intersect2DRCi = PointInCircle(ClosestPointOnRectangleFromPointRXY(Rectangle, aCircle.X, aCircle.Y), aCircle)
End Function
Public Function Intersect2DCiCi(Circle1 As TCircle, Circle2 As TCircle) As Boolean
  Intersect2DCiCi = (LayDistance2DXY(Circle1.X, Circle1.Y, Circle2.X, Circle2.Y) <= ((Circle1.Radius + Circle2.Radius) * (Circle1.Radius + Circle2.Radius)))
End Function
Public Function Intersect2DPg(Poly1 As TPolygon2D, Poly2 As TPolygon2D) As Boolean
Dim i As Long, j As Long, Poly1Trailer As Long, Poly2Trailer As Long
  Intersect2DPg = False
  If (UBound(Poly1.Arr) < 2) Or (UBound(Poly2.Arr) < 2) Then Exit Function
  Poly1Trailer = UBound(Poly1.Arr)
  For i = 0 To UBound(Poly1.Arr)
    Poly2Trailer = UBound(Poly2.Arr)
    For j = 0 To UBound(Poly2.Arr)
      If Intersect2DP(Poly1.Arr(i), Poly1.Arr(Poly1Trailer), Poly2.Arr(j), Poly2.Arr(Poly2Trailer)) Then
        Intersect2DPg = True
        Exit Function
      End If
      Poly2Trailer = j
    Next
    Poly1Trailer = i
  Next
End Function
Public Function Intersect3DXYfuz(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double, Optional fuzzy As Double = 0#) As Boolean
  Intersect3DXYfuz = (IsEqual(LayDistanceSegmentToSegment3D(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3, X4, Y4, z4), fuzzy))
End Function
Public Function Intersect3DPfuz(P1 As TPoint3D, P2 As TPoint3D, P3 As TPoint3D, P4 As TPoint3D, Optional fuzzy As Double = 0#) As Boolean
  Intersect3DPfuz = Intersect3DXYfuz(P1.X, P1.Y, P1.z, P2.X, P2.Y, P2.z, P3.X, P3.Y, P3.z, P4.X, P4.Y, P4.z, fuzzy)
End Function
Public Function Intersect3DSfuz(Segment1 As TSegment3D, Segment2 As TSegment3D, Optional fuzzy As Double = 0#) As Boolean
  Intersect3DSfuz = Intersect3DPfuz(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), fuzzy)
End Function
Public Function Intersect3DSSph(Segment As TSegment3D, Sphere As TSphere) As Boolean
Dim a As Double, b As Double, c As Double
  a = LayDistance3DS(Segment)
  b = 2 * ((Segment.P(1).X - Segment.P(0).X) * (Segment.P(0).X - Sphere.X) + (Segment.P(1).Y - Segment.P(0).Y) * (Segment.P(0).Y - Sphere.Y) + (Segment.P(1).z - Segment.P(0).z) * (Segment.P(0).z - Sphere.z))
  c = Sqr(Sphere.X) + Sqr(Sphere.Y) + Sqr(Sphere.z) + Sqr(Segment.P(0).X) + Sqr(Segment.P(0).Y) + Sqr(Segment.P(0).z) - 2 * (Sphere.X * Segment.P(0).X + Sphere.Y * Segment.P(0).Y + Sphere.z * Segment.P(0).z) - Sqr(Sphere.Radius)
  '//Result =((B * B - 4 * A * C) >= 0)
  Intersect3DSSph = GreaterThanOrEqual((b * b - 4 * a * c), 0#)
End Function
Public Function Intersect3DLToutP(Line As TLine3D, Triangle As TTriangle3D, IPoint As TPoint3D) As Boolean
Dim u As TVector3D, v As TVector3D, n As TVector3D, dir As TVector3D, w0 As TVector3D, W As TVector3D
Dim a As Double, b As Double, r As Double, uu As Double, uv As Double, vv As Double, wu As Double, wv As Double, D As Double, S As Double, t As Double
  
  Intersect3DLToutP = False
  
  u.X = Triangle.P(1).X - Triangle.P(0).X
  u.Y = Triangle.P(1).Y - Triangle.P(0).Y
  u.z = Triangle.P(1).z - Triangle.P(0).z
  
  v.X = Triangle.P(2).X - Triangle.P(0).X
  v.Y = Triangle.P(2).Y - Triangle.P(0).Y
  v.z = Triangle.P(2).z - Triangle.P(0).z
  
  n = Mul3D(u, v)
  
  If IsEqual(DotProduct3D(u, n), 0#) Then
    '(*
    '   The Triangle is degenerate, ie As  vertices are all
    '   collinear to each other and/or not unique.
    '*)
    Exit Function
  End If
  
  dir.X = Line.P(1).X - Line.P(0).X
  dir.Y = Line.P(1).Y - Line.P(0).Y
  dir.z = Line.P(1).z - Line.P(0).z
  
  w0.X = Line.P(0).X - Triangle.P(0).X
  w0.Y = Line.P(0).Y - Triangle.P(0).Y
  w0.z = Line.P(0).z - Triangle.P(0).z
  
  a = DotProduct3D(n, w0) * -1#
  b = DotProduct3D(n, dir)
  
  If IsEqual(Abs(b), 0#) Then
    Exit Function
    '(*
    '   A further test can be done to determine if the
    '   ray is coplanar to the Triangle.
    '   In any case the test for unique point intersection
    '   has failed.
    '   if IsEqual(a,0.0) then ray is coplanar to Triangle
    '*)
  End If
  
  r = a / b
  
  If IsEqual(r, 0#) Then
    Exit Function
  End If
  
  IPoint.X = Line.P(0).X + (r * dir.X)
  IPoint.Y = Line.P(0).Y + (r * dir.Y)
  IPoint.z = Line.P(0).z + (r * dir.z)
  
  W.X = IPoint.X - Triangle.P(0).X
  W.Y = IPoint.Y - Triangle.P(0).Y
  W.z = IPoint.z - Triangle.P(0).z
  
  uu = DotProduct3D(u, u)
  uv = DotProduct3D(u, v)
  vv = DotProduct3D(v, v)
  wu = DotProduct3D(W, u)
  wv = DotProduct3D(W, v)
  
  D = uv * uv - uu * vv
  
  '// get and test parametric coords
  S = ((uv * wv) - (vv * wu)) / D
  
  If (S < 0#) Or (S > 1#) Then
    '(* Intersection is outside of Triangle *)
    Exit Function
  End If
  
  t = ((uv * wu) - (uu * wv)) / D
  
  If (t < 0#) Or ((S + t) > 1#) Then
    '(* Intersection is outside of Triangle *)
    Exit Function
  End If
  
  Intersect3DLToutP = True
End Function
Public Function Intersect3DSph(Sphere1 As TSphere, Sphere2 As TSphere) As Boolean
  Intersect3DSph = (LayDistance3DXY(Sphere1.X, Sphere1.Y, Sphere1.z, Sphere2.X, Sphere2.Y, Sphere2.z) <= ((Sphere1.Radius + Sphere2.Radius) * (Sphere1.Radius + Sphere2.Radius)))
End Function
'(* End of Intersect *)

'SimpleIntersect
Public Function SimpleIntersectXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Boolean
  SimpleIntersectXY = (((Orientation2D(X1, Y1, X2, Y2, X3, Y3) * Orientation2D(X1, Y1, X2, Y2, X4, Y4)) <= 0) And _
                       ((Orientation2D(X3, Y3, X4, Y4, X1, Y1) * Orientation2D(X3, Y3, X4, Y4, X2, Y2)) <= 0))
End Function
Public Function SimpleIntersectP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D) As Boolean
  SimpleIntersectP = SimpleIntersectXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Point4.X, Point4.Y)
End Function
Public Function SimpleIntersectS(Segment1 As TSegment2D, Segment2 As TSegment2D) As Boolean
  SimpleIntersectS = SimpleIntersectP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1))
End Function
'(* End of SimpleIntersect *)

'ThickSegmentIntersect:
Public Function ThickSegmentIntersect2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, Thickness As Double) As Boolean
  ThickSegmentIntersect2DXY = LessThanOrEqual(Thickness * Thickness, LayDistanceSegmentToSegment2D(X1, Y1, X2, Y2, X3, Y3, X4, Y4))
End Function
Public Function ThickSegmentIntersect2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D, Thickness As Double) As Boolean
  ThickSegmentIntersect2DP = ThickSegmentIntersect2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Point4.X, Point4.Y, Thickness)
End Function
Public Function ThickSegmentIntersect2DS(Segment1 As TSegment2D, Segment2 As TSegment2D, Thickness As Double) As Boolean
  ThickSegmentIntersect2DS = ThickSegmentIntersect2DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Thickness)
End Function
Public Function ThickSegmentIntersect3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double, Thickness As Double) As Boolean
  ThickSegmentIntersect3DXY = LessThanOrEqual(Thickness * Thickness, LayDistanceSegmentToSegment3D(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3, X4, Y4, z4))
End Function
Public Function ThickSegmentIntersect3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Point4 As TPoint3D, Thickness As Double) As Boolean
  ThickSegmentIntersect3DP = ThickSegmentIntersect3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Point4.X, Point4.Y, Point4.z, Thickness)
End Function
Public Function ThickSegmentIntersect3DS(Segment1 As TSegment3D, Segment2 As TSegment3D, Thickness As Double) As Boolean
  ThickSegmentIntersect3DS = ThickSegmentIntersect3DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Thickness)
End Function
'(* End of ThickSegmentIntersect *)

'IntersectionPoint:
Public Sub IntersectionPoint2DXYN(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, nX As Double, nY As Double)
Dim Ratio As Double, Dx1 As Double, Dx2 As Double, Dx3 As Double, Dy1 As Double, Dy2 As Double, Dy3 As Double
  Dx1 = X2 - X1
  Dx2 = X4 - X3
  Dx3 = X1 - X3

  Dy1 = Y2 - Y1
  Dy2 = Y1 - Y3
  Dy3 = Y4 - Y3

  Ratio = Dx1 * Dy3 - Dy1 * Dx2
  If NotEqual(Ratio, 0#) Then
    Ratio = (Dy2 * Dx2 - Dx3 * Dy3) / Ratio
    nX = X1 + Ratio * Dx1
    nY = Y1 + Ratio * Dy1
  Else
    '//If Collinear(x1,y1,x2,y2,x3,y3) then
    If IsEqual((Dx1 * -Dy2), (-Dx3 * Dy1)) Then
      nX = X3
      nY = Y3
    Else
      nX = X4
      nY = Y4
    End If
  End If
End Sub
Public Sub IntersectionPoint2DPN(P1 As TPoint2D, P2 As TPoint2D, P3 As TPoint2D, P4 As TPoint2D, nX As Double, nY As Double)
  Call IntersectionPoint2DXYN(P1.X, P1.Y, P2.X, P2.Y, P3.X, P3.Y, P4.X, P4.Y, nX, nY)
End Sub
Public Function IntersectionPoint2DP(P1 As TPoint2D, P2 As TPoint2D, P3 As TPoint2D, P4 As TPoint2D) As TPoint2D
  Call IntersectionPoint2DXYN(P1.X, P1.Y, P2.X, P2.Y, P3.X, P3.Y, P4.X, P4.Y, IntersectionPoint2DP.X, IntersectionPoint2DP.Y)
End Function
Public Function IntersectionPoint2DS(Segment1 As TSegment2D, Segment2 As TSegment2D) As TPoint2D
  IntersectionPoint2DS = IntersectionPoint2DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1))
End Function
Public Function IntersectionPoint2DL(Line1 As TLine2D, Line2 As TLine2D) As TPoint2D
Dim Dx1 As Double, Dx2 As Double, Dx3 As Double, Dy1 As Double, Dy2 As Double, Dy3 As Double
Dim det As Double, Ratio As Double
  Dx1 = Line1.P(0).X - Line1.P(1).X
  Dx2 = Line2.P(0).X - Line2.P(1).X
  Dx3 = Line2.P(1).X - Line1.P(1).X
  Dy1 = Line1.P(0).Y - Line1.P(1).Y
  Dy2 = Line2.P(0).Y - Line2.P(1).Y
  Dy3 = Line2.P(1).Y - Line1.P(1).Y
  
  det = (Dx2 * Dy1) - (Dy2 * Dx1)
  
  If IsEqual(det, 0#) Then
    If IsEqual((Dx2 * Dy3), (Dy2 * Dx3)) Then
      IntersectionPoint2DL.X = Line2.P(0).X
      IntersectionPoint2DL.Y = Line2.P(0).Y
      Exit Function
    End If
  Else
    Exit Function
  End If
  
  Ratio = ((Dx1 * Dy3) - (Dy1 * Dx3)) / det
  
  IntersectionPoint2DL.X = (Ratio * Dx2) + Line2.P(2).X
  IntersectionPoint2DL.Y = (Ratio * Dy2) + Line2.P(2).Y
End Function
Public Sub IntersectionPoint2DCiP(Circle1 As TCircle, Circle2 As TCircle, Point1 As TPoint2D, Point2 As TPoint2D)
Dim Dist As Double, a As Double, H As Double, RatioA As Double, RatioH As Double
Dim Dx As Double, Dy As Double, Phi As TPoint2D, r1Sqr As Double, r2Sqr As Double, dstSqr As Double
  Dist = Distance2DXY(Circle1.X, Circle1.Y, Circle2.X, Circle2.Y)

  dstSqr = Dist * Dist
  r1Sqr = Circle1.Radius * Circle1.Radius
  r2Sqr = Circle2.Radius * Circle2.Radius

  a = (dstSqr - r2Sqr + r1Sqr) / (2 * Dist)
  H = Sqr(r1Sqr - (a * a))

  RatioA = a / Dist
  RatioH = H / Dist

  Dx = Circle2.X - Circle1.X
  Dy = Circle2.Y - Circle1.Y

  Phi.X = Circle1.X + (RatioA * Dx)
  Phi.Y = Circle1.Y + (RatioA * Dy)

  Dx = Dx * RatioH
  Dy = Dy * RatioH

  Point1.X = Phi.X + Dy
  Point1.Y = Phi.Y - Dx

  Point2.X = Phi.X - Dy
  Point2.Y = Phi.Y + Dx
End Sub
Public Sub IntersectionPoint2DST(Segment As TSegment2D, Triangle As TTriangle2D, ICnt As Long, I1 As TPoint2D, I2 As TPoint2D)
Dim ix As Double, iy As Double
  ICnt = 0
  If Intersect2DSXY(Segment, TriangleEdge2D(Triangle, 1), ix, iy) Then
    I1.X = ix
    I1.Y = iy
    Call Inc(ICnt)
  End If
  If Intersect2DSXY(Segment, TriangleEdge2D(Triangle, 2), ix, iy) Then
    If ICnt = 1 Then
      I2.X = ix
      I2.Y = iy
      Call Inc(ICnt)
      Exit Sub
    End If
    I1.X = ix
    I1.Y = iy
    Call Inc(ICnt)
  End If
  If Intersect2DSXY(Segment, TriangleEdge2D(Triangle, 3), ix, iy) Then
    If ICnt = 1 Then
      I2.X = ix
      I2.Y = iy
      Call Inc(ICnt)
      Exit Sub
    End If
    I1.X = ix
    I1.Y = iy
    Call Inc(ICnt)
  End If
End Sub
Public Sub IntersectionPoint3DLT(Line As TLine3D, Triangle As TTriangle3D, IPoint As TPoint3D)
Dim u As TVector3D, v As TVector3D, n As TVector3D, dir As TVector3D, w0  As TVector3D
Dim a As Double, b As Double, r As Double
  u.X = Triangle.P(1).X - Triangle.P(0).X
  u.Y = Triangle.P(1).Y - Triangle.P(0).Y
  u.z = Triangle.P(1).z - Triangle.P(0).z
  
  v.X = Triangle.P(2).X - Triangle.P(0).X
  v.Y = Triangle.P(2).Y - Triangle.P(0).Y
  v.z = Triangle.P(2).z - Triangle.P(0).z
  
  n = Mul3D(u, v)
  
  dir.X = Line.P(1).X - Line.P(0).X
  dir.Y = Line.P(1).Y - Line.P(0).Y
  dir.z = Line.P(1).z - Line.P(0).z
  
  w0.X = Line.P(0).X - Triangle.P(0).X
  w0.Y = Line.P(0).Y - Triangle.P(0).Y
  w0.z = Line.P(0).z - Triangle.P(0).z
  
  a = DotProduct3D(n, w0) * -1#
  b = DotProduct3D(n, dir)
  r = a / b
  
  IPoint.X = Line.P(0).X + (r * dir.X)
  IPoint.Y = Line.P(0).Y + (r * dir.Y)
  IPoint.z = Line.P(0).z + (r * dir.z)
End Sub
Public Sub IntersectionPoint2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, cX As Double, cY As Double, Radius As Double, ICnt As Long, Ix1 As Double, Iy1 As Double, Ix2 As Double, Iy2 As Double)
Dim Px As Double, Py As Double, S1In As Boolean, S2In As Boolean, H As Double, a As Double
  ICnt = 0
  S1In = PointInCircle2DXYR(X1, Y1, cX, cY, Radius)
  S2In = PointInCircle2DXYR(X2, Y2, cX, cY, Radius)
  If S1In And S2In Then
    ICnt = 2
    Ix1 = X1
    Iy1 = Y1
    Ix2 = X2
    Iy2 = Y2
    Exit Sub
  End If
  If S1In Or S2In Then
    ICnt = 2
    Call ClosestPointOnLineFromPointS2DXY(X1, Y1, X2, Y2, cX, cY, Px, Py)
    If S1In Then
      H = Distance2DXY(Px, Py, cX, cY)
      a = Sqr((Radius * Radius) - (H * H))
      Ix1 = X1
      Iy1 = Y1
      Call ProjectPointS2DSD(Px, Py, X2, Y2, a, Ix2, Iy2)
    ElseIf S2In Then
      H = Distance2DXY(Px, Py, cX, cY)
      a = Sqr((Radius * Radius) - (H * H))
      Ix1 = X2
      Iy1 = Y2
      Call ProjectPointS2DSD(Px, Py, X1, Y1, a, Ix2, Iy2)
    End If
    Exit Sub
  End If
  
  Call ClosestPointOnSegmentFromPointS2DXY(X1, Y1, X2, Y2, cX, cY, Px, Py)
  
  If (IsEqual(X1, Px) And IsEqual(Y1, Py)) Or _
     (IsEqual(X2, Px) And IsEqual(Y2, Py)) Then
    Exit Sub
  Else
    H = Distance2DXY(Px, Py, cX, cY)
    If H > Radius Then
      Exit Sub
    ElseIf IsEqual(H, Radius) Then
      ICnt = 1
      Ix1 = Px
      Iy1 = Py
      Exit Sub
    ElseIf IsEqual(H, 0#) Then
      ICnt = 2
      Call ProjectPointS2DSD(cX, cY, X1, Y1, Radius, Ix1, Iy1)
      Call ProjectPointS2DSD(cX, cY, X2, Y2, Radius, Ix2, Iy2)
      Exit Sub
    Else
      ICnt = 2
      a = Sqr((Radius * Radius) - (H * H))
      Call ProjectPointS2DSD(Px, Py, X1, Y1, a, Ix1, Iy1)
      Call ProjectPointS2DSD(Px, Py, X2, Y2, a, Ix2, Iy2)
      Exit Sub
    End If
  End If
End Sub
Public Sub IntersectionPoint2DSCi(Segment As TSegment2D, aCircle As TCircle, ICnt As Long, I1 As TPoint2D, I2 As TPoint2D)
  Call IntersectionPoint2DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, aCircle.X, aCircle.Y, aCircle.Radius, ICnt, I1.X, I1.Y, I2.X, I2.Y)
End Sub
'(* End of IntersectionPoint *)

'NormalizeAngle:
Public Function NormalizeAngle(Angle As Double) As Double
  NormalizeAngle = Angle
  If NormalizeAngle > 360# Then
    NormalizeAngle = NormalizeAngle - (Trunc(NormalizeAngle / 360#) * 360#)
  ElseIf NormalizeAngle < 0 Then
    Do While NormalizeAngle < 0#
      NormalizeAngle = NormalizeAngle + 360#
    Loop
  End If
End Function
'(* Normalize Angle *)

'VerticalMirror:
Public Function VerticalMirror(Angle As Double) As Double
  VerticalMirror = Angle
  If IsEqual(Angle, 0#) Or _
     IsEqual(Angle, 180#) Or _
     IsEqual(Angle, 360#) Then Exit Function
  VerticalMirror = 360 - VerticalMirror
End Function
'(* Vertical Mirror *)

'HorizontalMirror:
Public Function HorizontalMirror(Angle As Double) As Double
  HorizontalMirror = Angle
  If HorizontalMirror <= 180 Then
    HorizontalMirror = 180 - HorizontalMirror
  Else
    HorizontalMirror = 540 - HorizontalMirror
  End If
End Function
'(* Horizontal Mirror *)

'Quadrant:
Public Function QuadrantA(Angle As Double) As Long
  QuadrantA = 0
  If (Angle >= 0) And (Angle < 90) Then
    QuadrantA = 1
  ElseIf (Angle >= 90) And (Angle < 180) Then
    QuadrantA = 2
  ElseIf (Angle >= 180) And (Angle < 270) Then
    QuadrantA = 3
  ElseIf (Angle >= 270) And (Angle < 360) Then
    QuadrantA = 4
  ElseIf Angle = 360 Then
    QuadrantA = 1
  End If
End Function
Public Function Quadrant2DXY(X As Double, Y As Double) As Long
  If (X > 0#) And (Y >= 0#) Then
    Quadrant2DXY = 1
  ElseIf (X <= 0#) And (Y > 0#) Then
    Quadrant2DXY = 2
  ElseIf (X < 0#) And (Y <= 0#) Then
    Quadrant2DXY = 3
  ElseIf (X >= 0#) And (Y < 0#) Then
    Quadrant2DXY = 4
  Else
    Quadrant2DXY = 0
  End If
End Function
Public Function Quadrant2DP(Point As TPoint2D) As Long
  Quadrant2DP = Quadrant2DXY(Point.X, Point.Y)
End Function
'(* End of Quadrant *)

'VertexAngle:
Public Function VertexAngle2DXY(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal X3 As Double, ByVal Y3 As Double) As Double
Dim Dist As Double, InputTerm As Double
  '(*
  '   Using the cosine identity As
  '   cosA = (b ^ 2 + c ^ 2 - a ^ 2) / (2 * b * c)
  '   a = Cos   '((b^2 + c^2 - a^2) / (2*b*c))
  '
  '   Where As
  '
  '   a,b and c : are edges in the triangle
  '   A         : is the angle at the vertex opposite edge 'a'
  '               aka the edge defined by the vertex <x1y1-x2y2-x3y3>
  '
  '*)
  '(* Quantify coordinates *)
  X1 = X1 - X2
  X3 = X3 - X2
  Y1 = Y1 - Y2
  Y3 = Y3 - Y2
  
  '(* Calculate Ley Distance *)
  Dist = (X1 * X1 + Y1 * Y1) * (X3 * X3 + Y3 * Y3)
  
  If IsEqual(Dist, 0#) Then
    VertexAngle2DXY = 0#
  Else
    InputTerm = (X1 * X3 + Y1 * Y3) / Sqr(Dist)
    If IsEqual(InputTerm, 1#) Then
      VertexAngle2DXY = 0#
    ElseIf IsEqual(InputTerm, -1#) Then
      VertexAngle2DXY = 180#
    Else
      VertexAngle2DXY = ArcCos(InputTerm) * C180DivPI
    End If
  End If
End Function
Public Function VertexAngle2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Double
  VertexAngle2DP = VertexAngle2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function VertexAngle3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As Double
Dim Dist As Double
  '(*
  '   Method is the same as the one described in
  '   the above r, ine.
  '*)
  '(* Quantify coordinates *)
  X1 = X1 - X2
  X3 = X3 - X2
  Y1 = Y1 - Y2
  Y3 = Y3 - Y2
  z1 = z1 - z2
  z3 = z3 - z2
  '(* Calculate Ley Distance *)
  Dist = (X1 * X1 + Y1 * Y1 + z1 * z1) * (X3 * X3 + Y3 * Y3 + z3 * z3)
  If IsEqual(Dist, 0#) Then
    VertexAngle3DXY = 0#
  Else
    VertexAngle3DXY = ArcCos((X1 * X3 + Y1 * Y3 + z1 * z3) / Sqr(Dist)) * C180DivPI
  End If
End Function
Public Function VertexAngle3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As Double
  VertexAngle3DP = VertexAngle3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z)
End Function
'(* End of Vertex Angle *)

'OrientedVertexAngle:
Public Function OrientedVertexAngle2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Optional Orient As Long = Clockwise) As Double
  OrientedVertexAngle2DXY = VertexAngle2DXY(X1, Y1, X2, Y2, X3, Y3)
  If Orientation2D(X1, Y1, X2, Y2, X3, Y3) <> Orient Then
    OrientedVertexAngle2DXY = 360# - OrientedVertexAngle2DXY
  End If
End Function
Public Function OrientedVertexAngle2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Optional Orient As Long = Clockwise) As Double
  OrientedVertexAngle2DP = OrientedVertexAngle2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Orient)
End Function
'(* End of Oriented Vertex Angle *)

'CartesianAngle
Public Function CartesianAngle(X As Double, Y As Double) As Double
  If (X > 0#) And (Y > 0#) Then
    CartesianAngle = (Atn(Y / X) * C180DivPI)
  ElseIf (X < 0#) And (Y > 0#) Then
    CartesianAngle = (Atn(-X / Y) * C180DivPI) + 90#
  ElseIf (X < 0#) And (Y < 0#) Then
    CartesianAngle = (Atn(Y / X) * C180DivPI) + 180#
  ElseIf (X > 0#) And (Y < 0#) Then
    CartesianAngle = (Atn(-X / Y) * C180DivPI) + 270#
  ElseIf (X = 0#) And (Y > 0#) Then
    CartesianAngle = 90#
  ElseIf (X < 0#) And (Y = 0#) Then
    CartesianAngle = 180#
  ElseIf (X = 0#) And (Y < 0#) Then
    CartesianAngle = 270#
  Else
    CartesianAngle = 0#
  End If
End Function
Public Function CartesianAngleP(Point As TPoint2D) As Double
  CartesianAngleP = CartesianAngle(Point.X, Point.Y)
End Function
'(* End of Cartesian Angle *)

'RobustCartesianAngle
Public Function RobustCartesianAngle(X As Double, Y As Double, Optional Epsilon As Double = Epsilon_High) As Double
  If (X > 0#) And (Y > 0#) Then
    RobustCartesianAngle = (Atn(Y / X) * C180DivPI)
  ElseIf (X < 0#) And (Y > 0#) Then
    RobustCartesianAngle = (Atn(-X / Y) * C180DivPI) + 90#
  ElseIf (X < 0#) And (Y < 0#) Then
    RobustCartesianAngle = (Atn(Y / X) * C180DivPI) + 180#
  ElseIf (X > 0#) And (Y < 0#) Then
    RobustCartesianAngle = (Atn(-X / Y) * C180DivPI) + 270#
  ElseIf IsEqualEps(X, 0#, Epsilon) And (Y > 0#) Then
    RobustCartesianAngle = 90#
  ElseIf (X < 0#) And IsEqualEps(Y, 0#, Epsilon) Then
    RobustCartesianAngle = 180#
  ElseIf IsEqual(X, 0#) And (Y < 0#) Then
    RobustCartesianAngle = 270#
  Else
    RobustCartesianAngle = 0#
  End If
End Function
Public Function RobustCartesianAngleP(Point As TPoint2D, Optional Epsilon As Double = Epsilon_High) As Double
  RobustCartesianAngleP = RobustCartesianAngle(Point.X, Point.Y, Epsilon)
End Function
'(* End of Robust Cartesian Angle *)

'SegmentIntersectAngle:
Public Function SegmentIntersectAngle2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D) As Double
Dim TempPoint As TPoint2D
  SegmentIntersectAngle2DP = -1
  If Intersect2DP(Point1, Point2, Point3, Point4) Then
    TempPoint = IntersectionPoint2DP(Point1, Point2, Point3, Point4)
    SegmentIntersectAngle2DP = VertexAngle2DP(Point1, TempPoint, Point4)
  End If
End Function
Public Function SegmentIntersectAngle2DS(Segment1 As TSegment2D, Segment2 As TSegment2D) As Double
Dim TempPoint As TPoint2D
  SegmentIntersectAngle2DS = -1
  If Intersect2DS(Segment1, Segment2) Then
    TempPoint = IntersectionPoint2DS(Segment1, Segment2)
    SegmentIntersectAngle2DS = VertexAngle2DP(Segment1.P(0), TempPoint, Segment2.P(0))
  End If
End Function
'(* End of SegmentIntersectAngle *)

'InPortal
Public Function InPortal2D(P As TPoint2D) As Boolean
  InPortal2D = PointInRectanglePXY(P, MinimumX, MinimumY, MaximumX, MaximumY)
End Function
Public Function InPortal3D(P As TPoint3D) As Boolean
  InPortal3D = LessThanOrEqual(MinimumX, P.X) And LessThanOrEqual(MaximumZ, P.X) And _
               LessThanOrEqual(MinimumY, P.Y) And LessThanOrEqual(MaximumY, P.Y) And _
               LessThanOrEqual(MinimumZ, P.Y) And LessThanOrEqual(MaximumZ, P.Y)
End Function
'(* End of InPortal *)

'HighestPoint
Public Function HighestPoint2DPg(Polygon As TPolygon2D) As TPoint2D
Dim i As Long
  If UBound(Polygon.Arr) = 0 Then Exit Function
  HighestPoint2DPg = Polygon.Arr(0)
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).Y > HighestPoint2DPg.Y Then HighestPoint2DPg = Polygon.Arr(i)
  Next
End Function
Public Function HighestPoint2DPA(Point() As TPoint2D) As TPoint2D
Dim i As Long
  If UBound(Point) = 0 Then Exit Function
  HighestPoint2DPA = Point(0)
  For i = 1 To UBound(Point)
    If Point(i).Y > HighestPoint2DPA.Y Then HighestPoint2DPA = Point(i)
  Next
End Function
Public Function HighestPoint2DT(Triangle As TTriangle2D) As TPoint2D 'OK!!
  HighestPoint2DT = Triangle.P(0)
  If Triangle.P(1).Y > HighestPoint2DT.Y Then HighestPoint2DT = Triangle.P(1)
  If Triangle.P(2).Y > HighestPoint2DT.Y Then HighestPoint2DT = Triangle.P(2)
End Function
Public Function HighestPoint3DT(Triangle As TTriangle3D) As TPoint3D
  HighestPoint3DT = Triangle.P(0)
  If Triangle.P(1).Y > HighestPoint3DT.Y Then HighestPoint3DT = Triangle.P(1)
  If Triangle.P(2).Y > HighestPoint3DT.Y Then HighestPoint3DT = Triangle.P(2)
End Function
Public Function HighestPoint2DQ(Quadix As TQuadix2D) As TPoint2D
  HighestPoint2DQ = Quadix.P(0)
  If Quadix.P(1).Y > HighestPoint2DQ.Y Then HighestPoint2DQ = Quadix.P(1)
  If Quadix.P(2).Y > HighestPoint2DQ.Y Then HighestPoint2DQ = Quadix.P(2)
  If Quadix.P(3).Y > HighestPoint2DQ.Y Then HighestPoint2DQ = Quadix.P(3)
End Function
Public Function HighestPoint3DQ(Quadix As TQuadix3D) As TPoint3D
  HighestPoint3DQ = Quadix.P(0)
  If Quadix.P(1).Y > HighestPoint3DQ.Y Then HighestPoint3DQ = Quadix.P(1)
  If Quadix.P(2).Y > HighestPoint3DQ.Y Then HighestPoint3DQ = Quadix.P(2)
  If Quadix.P(3).Y > HighestPoint3DQ.Y Then HighestPoint3DQ = Quadix.P(3)
End Function
'(* End of HighestPoint *)

'LowestPoint
Public Function LowestPoint2DPg(Polygon As TPolygon2D) As TPoint2D
Dim i As Long
  If UBound(Polygon.Arr) = 0 Then Exit Function
  LowestPoint2DPg = Polygon.Arr(0)
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).Y < LowestPoint2DPg.Y Then LowestPoint2DPg = Polygon.Arr(i)
  Next
End Function
Public Function LowestPoint2DPA(Point() As TPoint2D) As TPoint2D
Dim i As Long
  If UBound(Point) = 0 Then Exit Function
  LowestPoint2DPA = Point(0)
  For i = 1 To UBound(Point)
    If Point(i).Y < LowestPoint2DPA.Y Then LowestPoint2DPA = Point(i)
  Next
End Function
Public Function LowestPoint2DT(Triangle As TTriangle2D) As TPoint2D
  LowestPoint2DT = Triangle.P(0)
  If Triangle.P(1).Y < LowestPoint2DT.Y Then LowestPoint2DT = Triangle.P(1)
  If Triangle.P(2).Y < LowestPoint2DT.Y Then LowestPoint2DT = Triangle.P(2)
End Function
Public Function LowestPoint3DT(Triangle As TTriangle3D) As TPoint3D
  LowestPoint3DT = Triangle.P(0)
  If Triangle.P(1).Y < LowestPoint3DT.Y Then LowestPoint3DT = Triangle.P(1)
  If Triangle.P(2).Y < LowestPoint3DT.Y Then LowestPoint3DT = Triangle.P(2)
End Function
Public Function LowestPoint2DQ(Quadix As TQuadix2D) As TPoint2D
  LowestPoint2DQ = Quadix.P(0)
  If Quadix.P(1).Y < LowestPoint2DQ.Y Then LowestPoint2DQ = Quadix.P(1)
  If Quadix.P(2).Y < LowestPoint2DQ.Y Then LowestPoint2DQ = Quadix.P(2)
  If Quadix.P(3).Y < LowestPoint2DQ.Y Then LowestPoint2DQ = Quadix.P(3)
End Function
Public Function LowestPoint3DQ(Quadix As TQuadix3D) As TPoint3D
  LowestPoint3DQ = Quadix.P(0)
  If Quadix.P(1).Y < LowestPoint3DQ.Y Then LowestPoint3DQ = Quadix.P(1)
  If Quadix.P(2).Y < LowestPoint3DQ.Y Then LowestPoint3DQ = Quadix.P(2)
  If Quadix.P(3).Y < LowestPoint3DQ.Y Then LowestPoint3DQ = Quadix.P(3)
End Function
'(* End of LowestPoint *)


'MostLeftPoint
Public Function MostLeftPointPg(Polygon As TPolygon2D) As TPoint2D
Dim i As Long
  If UBound(Polygon.Arr) = 0 Then Exit Function
  MostLeftPointPg = Polygon.Arr(0)
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).X < MostLeftPointPg.X Then MostLeftPointPg = Polygon.Arr(i)
  Next
End Function
Public Function MostLeftPointPA(Points() As TPoint2D) As TPoint2D
Dim i As Long
  If UBound(Points) = 0 Then Exit Function
  MostLeftPointPA = Points(0)
  For i = 1 To UBound(Points)
    If Points(i).X < MostLeftPointPA.X Then MostLeftPointPA = Points(i)
  Next
End Function
'(* End of Most Left Point *)

'MostRightPoint
Public Function MostRightPointPg(Polygon As TPolygon2D) As TPoint2D
Dim i As Long
  If UBound(Polygon.Arr) = 0 Then Exit Function
  MostRightPointPg = Polygon.Arr(0)
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).X < MostRightPointPg.X Then MostRightPointPg = Polygon.Arr(i)
  Next
End Function
Public Function MostRightPointPA(Points() As TPoint2D) As TPoint2D
Dim i As Long
  If UBound(Points) = 0 Then Exit Function
  MostRightPointPA = Points(0)
  For i = 1 To UBound(Points)
    If Points(i).X < MostRightPointPA.X Then MostRightPointPA = Points(i)
  Next
End Function
'(* End of Most Right Point *)

'MostUpperRight
Public Function MostUpperRightPg(Polygon As TPolygon2D) As TPoint2D
Dim i As Long
  If UBound(Polygon.Arr) = 0 Then Exit Function
  MostUpperRightPg = Polygon.Arr(0)
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).Y > MostUpperRightPg.Y Then
      MostUpperRightPg = Polygon.Arr(i)
    ElseIf Polygon.Arr(i).Y = MostUpperRightPg.Y Then
      If Polygon.Arr(i).X > MostUpperRightPg.X Then MostUpperRightPg = Polygon.Arr(i)
    End If
  Next
End Function
Public Function MostUpperRightPA(Point() As TPoint2D) As TPoint2D
Dim i As Long
  If UBound(Point) = 0 Then Exit Function
  MostUpperRightPA = Point(0)
  For i = 1 To UBound(Point)
    If Point(i).Y > MostUpperRightPA.Y Then
      MostUpperRightPA = Point(i)
    ElseIf Point(i).Y = MostUpperRightPA.Y Then
      If Point(i).X > MostUpperRightPA.X Then MostUpperRightPA = Point(i)
    End If
  Next
End Function
'(* End of Most Upper Right *)

'MostUpperLeft
Public Function MostUpperLeftPg(Polygon As TPolygon2D) As TPoint2D
Dim i As Long
  If UBound(Polygon.Arr) = 0 Then Exit Function
  MostUpperLeftPg = Polygon.Arr(0)
  For i = 1 To UBound(Polygon.Arr)
   If Polygon.Arr(i).Y > MostUpperLeftPg.Y Then
     MostUpperLeftPg = Polygon.Arr(i)
   ElseIf Polygon.Arr(i).Y = MostUpperLeftPg.Y Then
     If Polygon.Arr(i).X < MostUpperLeftPg.X Then
       MostUpperLeftPg = Polygon.Arr(i)
     End If
   End If
 Next
End Function
Public Function MostUpperLeftPA(Point() As TPoint2D) As TPoint2D
Dim i As Long
  If UBound(Point) = 0 Then Exit Function
  MostUpperLeftPA = Point(0)
  For i = 1 To UBound(Point)
   If Point(i).Y > MostUpperLeftPA.Y Then
     MostUpperLeftPA = Point(i)
   ElseIf Point(i).Y = MostUpperLeftPA.Y Then
     If Point(i).X < MostUpperLeftPA.X Then
       MostUpperLeftPA = Point(i)
     End If
   End If
 Next
End Function
'(* End of Most Upper Left *)

'MostLowerRight
Public Function MostLowerRightPg(Polygon As TPolygon2D) As TPoint2D
Dim i As Long
  If UBound(Polygon.Arr) = 0 Then Exit Function
  MostLowerRightPg = Polygon.Arr(0)
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).Y < MostLowerRightPg.Y Then
      MostLowerRightPg = Polygon.Arr(i)
    ElseIf Polygon.Arr(i).Y = MostLowerRightPg.Y Then
      If Polygon.Arr(i).X > MostLowerRightPg.X Then MostLowerRightPg = Polygon.Arr(i)
    End If
  Next
End Function
Public Function MostLowerRightPA(Point() As TPoint2D) As TPoint2D
Dim i As Long
  If UBound(Point) = 0 Then Exit Function
  MostLowerRightPA = Point(0)
  For i = 1 To UBound(Point)
    If Point(i).Y < MostLowerRightPA.Y Then
      MostLowerRightPA = Point(i)
    ElseIf Point(i).Y = MostLowerRightPA.Y Then
      If Point(i).X > MostLowerRightPA.X Then MostLowerRightPA = Point(i)
    End If
  Next
End Function
'(* End of Most Lower Right *)

'MostLowerLeft
Public Function MostLowerLeftPg(Polygon As TPolygon2D) As TPoint2D
Dim i As Long
  If UBound(Polygon.Arr) = 0 Then Exit Function
  MostLowerLeftPg = Polygon.Arr(0)
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).Y < MostLowerLeftPg.Y Then
      MostLowerLeftPg = Polygon.Arr(i)
    ElseIf Polygon.Arr(i).Y = MostLowerLeftPg.Y Then
       If Polygon.Arr(i).X < MostLowerLeftPg.X Then MostLowerLeftPg = Polygon.Arr(i)
    End If
  Next
End Function
Public Function MostLowerLeftPA(Point() As TPoint2D) As TPoint2D
Dim i As Long
  If UBound(Point) = 0 Then Exit Function
  MostLowerLeftPA = Point(0)
  For i = 1 To UBound(Point)
    If Point(i).Y < MostLowerLeftPA.Y Then
      MostLowerLeftPA = Point(i)
    ElseIf Point(i).Y = MostLowerLeftPA.Y Then
       If Point(i).X < MostLowerLeftPA.X Then MostLowerLeftPA = Point(i)
    End If
  Next
End Function
'(* End of Most Lower Left *)

'Min
Public Function Min2DP(Point1 As TPoint2D, Point2 As TPoint2D) As TPoint2D
  If Point1.X < Point2.X Then
    Min2DP = Point1
  ElseIf Point2.X < Point1.X Then
    Min2DP = Point2
  ElseIf Point1.Y < Point2.Y Then
    Min2DP = Point1
  Else
    Min2DP = Point2
  End If
End Function
Public Function Min3DP(Point1 As TPoint3D, Point2 As TPoint3D) As TPoint3D
  If Point1.X < Point2.X Then
    Min3DP = Point1
  ElseIf Point2.X < Point1.X Then
    Min3DP = Point2
  ElseIf Point1.Y < Point2.Y Then
    Min3DP = Point1
  ElseIf Point2.Y < Point1.Y Then
    Min3DP = Point2
  ElseIf Point1.z < Point2.z Then
    Min3DP = Point1
  Else
    Min3DP = Point2
  End If
End Function
'(* End of Minimum Between 2 Points *)

'Max
Public Function Max2DP(Point1 As TPoint2D, Point2 As TPoint2D) As TPoint2D
  If Point1.X > Point2.X Then
    Max2DP = Point1
  ElseIf Point2.X > Point1.X Then
    Max2DP = Point2
  ElseIf Point1.Y > Point2.Y Then
    Max2DP = Point1
  Else
    Max2DP = Point2
  End If
End Function
Public Function Max3DP(Point1 As TPoint3D, Point2 As TPoint3D) As TPoint3D
  If Point1.X > Point2.X Then
    Max3DP = Point1
  ElseIf Point2.X > Point1.X Then
    Max3DP = Point2
  ElseIf Point1.Y > Point2.Y Then
    Max3DP = Point1
  ElseIf Point2.Y > Point1.Y Then
    Max3DP = Point2
  ElseIf Point1.z > Point2.z Then
    Max3DP = Point1
  Else
    Max3DP = Point2
  End If
End Function
'(* End of Maximum Between 2 Points *)

'Coincident
Public Function Coincident2DP(Point1 As TPoint2D, Point2 As TPoint2D) As Boolean
  Coincident2DP = IsEqual2DP(Point1, Point2)
End Function
Public Function Coincident3DP(Point1 As TPoint3D, Point2 As TPoint3D) As Boolean
  Coincident3DP = IsEqual3DP(Point1, Point2)
End Function
Public Function Coincident2DS(Segment1 As TSegment2D, Segment2 As TSegment2D) As Boolean 'OK!!
  Coincident2DS = (Coincident2DP(Segment1.P(0), Segment2.P(0)) And Coincident2DP(Segment1.P(1), Segment2.P(1))) Or _
                  (Coincident2DP(Segment1.P(0), Segment2.P(1)) And Coincident2DP(Segment1.P(1), Segment2.P(0)))
End Function
Public Function Coincident3DS(Segment1 As TSegment3D, Segment2 As TSegment3D) As Boolean 'OK!!
  Coincident3DS = (Coincident3DP(Segment1.P(0), Segment2.P(0)) And Coincident3DP(Segment1.P(1), Segment2.P(1))) Or _
                  (Coincident3DP(Segment1.P(0), Segment2.P(1)) And Coincident3DP(Segment1.P(1), Segment2.P(0)))
End Function
Public Function Coincident2DT(Triangle1 As TTriangle2D, Triangle2 As TTriangle2D) As Boolean
ReDim Flag(0 To 2) As Boolean 'OK!!
Dim Count As Long, i As Long, j As Long
  Count = 0
  For i = 0 To 2: Flag(i) = False: Next
  For i = 0 To 2
    For j = 0 To 2
      If Not Flag(i) Then
        If Coincident2DP(Triangle1.P(i), Triangle2.P(j)) Then
          Call Inc(Count)
          Flag(i) = True
          Exit For
        End If
      End If
    Next
  Next
  Coincident2DT = (Count = 3)
End Function
Public Function Coincident3DT(Triangle1 As TTriangle3D, Triangle2 As TTriangle3D) As Boolean
ReDim Flag(0 To 2) As Boolean 'OK!!
Dim Count As Long, i As Long, j As Long
  Count = 0
  For i = 0 To 2: Flag(i) = False: Next
  For i = 0 To 2
    For j = 0 To 2
      If Not Flag(i) Then
        If Coincident3DP(Triangle1.P(i), Triangle2.P(j)) Then
          Call Inc(Count)
          Flag(i) = True
          Exit Function
        End If
      End If
    Next
  Next
  Coincident3DT = (Count = 3)
End Function
Public Function Coincident2DR(Rect1 As TRectangle, Rect2 As TRectangle) As Boolean
  Coincident2DR = Coincident2DP(Rect1.P(1), Rect2.P(1)) And Coincident2DP(Rect1.P(2), Rect2.P(2))
End Function
Public Function Coincident2DQ(Quad1 As TQuadix2D, Quad2 As TQuadix2D) As Boolean
ReDim Flag(0 To 3) As Boolean 'OK!!
Dim Count As Long, i As Long, j As Long
  Coincident2DQ = False
  If ConvexQuadix(Quad1) <> ConvexQuadix(Quad2) Then Exit Function
  Count = 0
  For i = 0 To 3: Flag(i) = False: Next
  For i = 0 To 3
    For j = 0 To 3
      If Not Flag(i) Then
        If Coincident2DP(Quad1.P(i), Quad2.P(j)) Then
          Call Inc(Count)
          Flag(i) = True
          Exit For
        End If
      End If
    Next
  Next
  Coincident2DQ = (Count = 4)
End Function
Public Function Coincident3DQ(Quad1 As TQuadix3D, Quad2 As TQuadix3D) As Boolean
  '(* to be implemented at a later date *)
  Coincident3DQ = False
End Function
Public Function Coincident2DCi(Circle1 As TCircle, Circle2 As TCircle) As Boolean 'OK!!
  Coincident2DCi = IsEqual(Circle1.X, Circle2.X) And _
                   IsEqual(Circle1.Y, Circle2.Y) And _
                   IsEqual(Circle1.Radius, Circle2.Radius)
End Function
Public Function CoincidentSph(Sphr1 As TSphere, Sphr2 As TSphere) As Boolean 'OK!!
  CoincidentSph = IsEqual(Sphr1.X, Sphr2.X) And _
                  IsEqual(Sphr1.Y, Sphr2.Y) And _
                  IsEqual(Sphr1.z, Sphr2.z) And _
                  IsEqual(Sphr1.Radius, Sphr2.Radius)
End Function
Public Function CoincidentO(Obj1 As TGeometricObject, Obj2 As TGeometricObject) As Boolean
  If (Obj1.ObjectType = geoPoint2D) And (Obj2.ObjectType = geoPoint2D) Then
    CoincidentO = Coincident2DP(Obj1.Point2D, Obj2.Point2D)
  ElseIf (Obj1.ObjectType = geoPoint3D) And (Obj2.ObjectType = geoPoint3D) Then
    CoincidentO = Coincident3DP(Obj1.Point3D, Obj2.Point3D)
  ElseIf (Obj1.ObjectType = geoSegment2D) And (Obj2.ObjectType = geoSegment2D) Then
    CoincidentO = Coincident2DS(Obj1.Segment2D, Obj2.Segment2D)
  ElseIf (Obj1.ObjectType = geoSegment3D) And (Obj2.ObjectType = geoSegment3D) Then
    CoincidentO = Coincident3DS(Obj1.Segment3D, Obj2.Segment3D)
  ElseIf (Obj1.ObjectType = geoTriangle2D) And (Obj2.ObjectType = geoTriangle2D) Then
    CoincidentO = Coincident2DT(Obj1.Triangle2D, Obj2.Triangle2D)
  ElseIf (Obj1.ObjectType = geoTriangle3D) And (Obj2.ObjectType = geoTriangle3D) Then
    CoincidentO = Coincident3DT(Obj1.Triangle3D, Obj2.Triangle3D)
  ElseIf (Obj1.ObjectType = geoQuadix2D) And (Obj2.ObjectType = geoQuadix2D) Then
    CoincidentO = Coincident2DQ(Obj1.Quadix2D, Obj2.Quadix2D)
  ElseIf (Obj1.ObjectType = geoQuadix3D) And (Obj2.ObjectType = geoQuadix3D) Then
    CoincidentO = Coincident3DQ(Obj1.Quadix3D, Obj2.Quadix3D)
  ElseIf (Obj1.ObjectType = geoCircle) And (Obj2.ObjectType = geoCircle) Then
    CoincidentO = Coincident2DCi(Obj1.aCircle, Obj2.aCircle)
  ElseIf (Obj1.ObjectType = geoSphere) And (Obj2.ObjectType = geoSphere) Then
    CoincidentO = CoincidentSph(Obj1.Sphere, Obj2.Sphere)
  Else
    CoincidentO = False
  End If
End Function
'(* End of Coincident  *)

'Parallel:
Public Function Parallel2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Parallel2DXY = IsEqualEps(((Y1 - Y2) * (X3 - X4)), ((Y3 - Y4) * (X1 - X2)), Epsilon)
End Function
Public Function Parallel2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Parallel2DP = Parallel2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Point4.X, Point4.Y, Epsilon)
End Function
Public Function Parallel2DS(Segment1 As TSegment2D, Segment2 As TSegment2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Parallel2DS = Parallel2DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Epsilon)
End Function
Public Function Parallel2DL(Line1 As TLine2D, Line2 As TLine2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Parallel2DL = Parallel2DP(Line1.P(0), Line1.P(1), Line2.P(0), Line2.P(1), Epsilon)
End Function
Public Function Parallel3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double, Optional Epsilon As Double = Epsilon_Medium) As Boolean
Dim Dx1 As Double, Dx2 As Double, Dy1 As Double, Dy2 As Double, Dz1 As Double, Dz2 As Double
  '(*
  '   Theory:
  '   If the gradients in the following planes x-y, y-z, z-x are equal then
  '   it can be said that the segments are parallel in 3D - ~ i think ~
  '   Worst case scenario As  6 floating point multiplications and 9 floating
  '   Point subtractions
  '*)
  Parallel3DXY = False
  
  Dx1 = X1 - X2
  Dx2 = X3 - X4
  
  Dy1 = Y1 - Y2
  Dy2 = Y3 - Y4
  
  Dz1 = z1 - z2
  Dz2 = z3 - z4
  
  If NotEqualEps((Dy1 * Dx2), (Dy2 * Dx1), Epsilon) Then Exit Function
  If NotEqualEps((Dz1 * Dy2), (Dz2 * Dy1), Epsilon) Then Exit Function
  If NotEqualEps((Dx1 * Dz2), (Dx2 * Dz1), Epsilon) Then Exit Function
  
  Parallel3DXY = True
End Function
Public Function Parallel3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Point4 As TPoint3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Parallel3DP = Parallel3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Point4.X, Point4.Y, Point4.z, Epsilon)
End Function
Public Function Parallel3DS(Segment1 As TSegment3D, Segment2 As TSegment3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Parallel3DS = Parallel3DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Epsilon)
End Function
Public Function Parallel3DL(Line1 As TLine3D, Line2 As TLine3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Parallel3DL = Parallel3DP(Line1.P(0), Line1.P(1), Line2.P(0), Line2.P(1), Epsilon)
End Function
'(* End of Parallel *)

'RobustParallel:
Public Function RobustParallel2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, Optional Epsilon As Double = Epsilon_Medium) As Boolean
Dim Px1 As Double, Py1 As Double, Px2 As Double, Py2 As Double
  Call ClosestPointOnLineFromPointS2DXY(X1, Y1, X2, Y2, X3, Y3, Px1, Py1)
  Call ClosestPointOnLineFromPointS2DXY(X1, Y1, X2, Y2, X4, Y4, Px2, Py2)
  RobustParallel2DXY = IsEqualEps(Distance2DXY(X3, Y3, Px1, Py1), Distance2DXY(X4, Y4, Px2, Py2), Epsilon)
End Function
Public Function RobustParallel2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustParallel2DP = RobustParallel2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Point4.X, Point4.Y, Epsilon)
End Function
Public Function RobustParallel2DS(Segment1 As TSegment2D, Segment2 As TSegment2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustParallel2DS = RobustParallel2DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Epsilon)
End Function
Public Function RobustParallel2DL(Line1 As TLine2D, Line2 As TLine2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustParallel2DL = RobustParallel2DP(Line1.P(0), Line1.P(1), Line2.P(0), Line2.P(1), Epsilon)
End Function
Public Function RobustParallel3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double, Optional Epsilon As Double = Epsilon_Medium) As Boolean
Dim Px1 As Double, Py1 As Double, Pz1 As Double, Px2 As Double, Py2 As Double, Pz2 As Double
  Call ClosestPointOnLineFromPointS3DXY(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3, Px1, Py1, Pz1)
  Call ClosestPointOnLineFromPointS3DXY(X1, Y1, z1, X2, Y2, z2, X4, Y4, z4, Px2, Py2, Pz2)
  RobustParallel3DXY = IsEqualEps(Distance3DXY(X3, Y3, z3, Px1, Py1, Pz1), Distance3DXY(X4, Y4, z4, Px2, Py2, Pz2), Epsilon)
End Function
Public Function RobustParallel3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Point4 As TPoint3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustParallel3DP = RobustParallel3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Point4.X, Point4.Y, Point3.z, Epsilon)
End Function
Public Function RobustParallel3DS(Segment1 As TSegment3D, Segment2 As TSegment3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustParallel3DS = RobustParallel3DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Epsilon)
End Function
Public Function RobustParallel3DL(Line1 As TLine3D, Line2 As TLine3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustParallel3DL = RobustParallel3DP(Line1.P(0), Line1.P(1), Line2.P(0), Line2.P(1), Epsilon)
End Function
'(* End of Robust Parallel *)

'Perpendicular
Public Function Perpendicular2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Perpendicular2DXY = IsEqualEps(-1# * (Y2 - Y1) * (Y4 - Y3), (X4 - X3) * (X2 - X1), Epsilon)
End Function
Public Function Perpendicular2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Perpendicular2DP = Perpendicular2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Point4.X, Point4.Y, Epsilon)
End Function
Public Function Perpendicular2DS(Segment1 As TSegment2D, Segment2 As TSegment2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Perpendicular2DS = Perpendicular2DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Epsilon)
End Function
Public Function Perpendicular2DL(Line1 As TLine2D, Line2 As TLine2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Perpendicular2DL = Perpendicular2DP(Line1.P(0), Line1.P(1), Line2.P(0), Line2.P(1), Epsilon)
End Function
Public Function Perpendicular3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double, Optional Epsilon As Double = Epsilon_Medium) As Boolean
Dim Dx1 As Double, Dx2 As Double, Dy1 As Double, Dy2 As Double, Dz1 As Double, Dz2 As Double
 '(*
 '   The dot product of the vector forms of the segments will
 '   be 0 If the segments are perpendicular
 '*)
  
  Dx1 = X1 - X2
  Dx2 = X3 - X4
  
  Dy1 = Y1 - Y2
  Dy2 = Y3 - Y4
  
  Dz1 = z1 - z2
  Dz2 = z3 - z4
  
  Perpendicular3DXY = IsEqualEps((Dx1 * Dx2) + (Dy1 * Dy2) + (Dz1 * Dz2), 0#, Epsilon)
End Function
Public Function Perpendicular3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Point4 As TPoint3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Perpendicular3DP = Perpendicular3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Point4.X, Point4.Y, Point4.z)
End Function
Public Function Perpendicular3DS(Segment1 As TSegment3D, Segment2 As TSegment3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Perpendicular3DS = Perpendicular3DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Epsilon)
End Function
Public Function Perpendicular3DL(Line1 As TLine3D, Line2 As TLine3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  Perpendicular3DL = Perpendicular3DP(Line1.P(0), Line1.P(1), Line2.P(0), Line2.P(1), Epsilon)
End Function
'(* End of Perpendicular *)

'RobustPerpendicular:
Public Function RobustPerpendicular2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, Optional Epsilon As Double = Epsilon_Medium) As Boolean
Dim P1x As Double, P1y As Double, P2x As Double, P2y As Double
  Call ClosestPointOnLineFromPointS2DXY(X1, Y1, X2, Y2, X3, Y3, P1x, P1y)
  Call ClosestPointOnLineFromPointS2DXY(X1, Y1, X2, Y2, X4, Y4, P2x, P2y)
  RobustPerpendicular2DXY = IsEqualEps(Distance2DXY(P1x, P1y, P2x, P2y), 0#, Epsilon)
End Function
Public Function RobustPerpendicular2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustPerpendicular2DP = RobustPerpendicular2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Point4.X, Point4.Y, Epsilon)
End Function
Public Function RobustPerpendicular2DS(Segment1 As TSegment2D, Segment2 As TSegment2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustPerpendicular2DS = RobustPerpendicular2DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Epsilon)
End Function
Public Function RobustPerpendicular2DL(Line1 As TLine2D, Line2 As TLine2D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustPerpendicular2DL = RobustPerpendicular2DP(Line1.P(0), Line1.P(1), Line2.P(0), Line2.P(1), Epsilon)
End Function
Public Function RobustPerpendicular3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double, Optional Epsilon As Double = Epsilon_Medium) As Boolean
Dim P1x As Double, P1y As Double, P1z As Double, P2x As Double, P2y As Double, P2z As Double
  Call ClosestPointOnLineFromPointS3DXY(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3, P1x, P1y, P1z)
  Call ClosestPointOnLineFromPointS3DXY(X1, Y1, z1, X2, Y2, z2, X4, Y4, z3, P2x, P2y, P2z)
  RobustPerpendicular3DXY = IsEqualEps(Distance3DXY(P1x, P1y, P1z, P2x, P2y, P2z), 0#, Epsilon)
End Function
Public Function RobustPerpendicular3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Point4 As TPoint3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustPerpendicular3DP = RobustPerpendicular3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Point4.X, Point4.Y, Point4.z, Epsilon)
End Function
Public Function RobustPerpendicular3DS(Segment1 As TSegment3D, Segment2 As TSegment3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustPerpendicular3DS = RobustPerpendicular3DP(Segment1.P(0), Segment1.P(1), Segment2.P(0), Segment2.P(1), Epsilon)
End Function
Public Function RobustPerpendicular3DL(Line1 As TLine3D, Line2 As TLine3D, Optional Epsilon As Double = Epsilon_Medium) As Boolean
  RobustPerpendicular3DL = RobustPerpendicular3DP(Line1.P(0), Line1.P(1), Line2.P(0), Line2.P(1), Epsilon)
End Function
'(* End of Robust Perpendicular *)

'LineToLineIntersect
Public Function LineToLineIntersect2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Boolean
  LineToLineIntersect2DXY = False
  If NotEqual((X1 - X2) * (Y3 - Y4), (Y1 - Y2) * (X3 - X4)) Then
    LineToLineIntersect2DXY = True
  ElseIf Collinear2D(X1, Y1, X2, Y2, X3, Y3) Then
    LineToLineIntersect2DXY = True
  End If
End Function
Public Function LineToLineIntersect2DL(Line1 As TLine2D, Line2 As TLine2D) As Boolean
  LineToLineIntersect2DL = LineToLineIntersect2DXY(Line1.P(0).X, Line1.P(0).Y, Line1.P(1).X, Line1.P(1).Y, Line2.P(0).X, Line2.P(0).Y, Line2.P(1).X, Line2.P(1).Y)
End Function
'(* Endof LineToLineIntersect *)

'RectangleToRectangleIntersect
Public Function RectangleToRectangleIntersectXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Boolean
  '(*
  '  Assumes that :   x1 < x2, y1 < y2, x3 < x4, y3 < y4
  '*)
  RectangleToRectangleIntersectXY = (X1 <= X4) And (X2 >= X3) And (Y1 <= Y4) And (Y2 >= Y3)
End Function
Public Function RectangleToRectangleIntersectR(Rectangle1 As TRectangle, Rectangle2 As TRectangle) As Boolean
  RectangleToRectangleIntersectR = RectangleToRectangleIntersectXY(Rectangle1.P(0).X, Rectangle1.P(0).Y, Rectangle1.P(2).X, Rectangle1.P(1).Y, _
                                                                   Rectangle2.P(0).X, Rectangle2.P(0).Y, Rectangle2.P(2).X, Rectangle2.P(1).Y)
End Function
'(* Endof Rectangle To Rectangle Intersect *)

'RectangleWithinRectangleIntersect
Public Function RectangleWithinRectangleXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Boolean
  RectangleWithinRectangleXY = PointInRectangleXY(X1, Y1, X3, Y3, X4, Y4) And PointInRectangleXY(X2, Y2, X3, Y3, X4, Y4)
End Function
Public Function RectangleWithinRectangle(Rectangle1 As TRectangle, Rectangle2 As TRectangle) As Boolean
  RectangleWithinRectangle = RectangleWithinRectangleXY(Rectangle1.P(0).X, Rectangle1.P(0).Y, Rectangle1.P(1).X, Rectangle1.P(1).Y, _
                                                        Rectangle2.P(0).X, Rectangle2.P(0).Y, Rectangle2.P(1).X, Rectangle2.P(1).Y)
End Function
'(* Endof Rectangle Within Rectangle Intersect *)

'CircleWithinRectangle
Public Function CircleWithinRectangleXY(X As Double, Y As Double, Radius As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Boolean
  CircleWithinRectangleXY = RectangleWithinRectangle(AABB2DCi(EquateCircleXY(X, Y, Radius)), EquateRectangleXY(X1, Y1, X2, Y2))
End Function
Public Function CircleWithinRectangle(aCircle As TCircle, aRectangle As TRectangle) As Boolean
  CircleWithinRectangle = RectangleWithinRectangle(AABB2DCi(aCircle), aRectangle)
End Function
'(* Endof Circle Within Rectangle Intersect *)

'TriangleWithinRectangle
Public Function TriangleWithinRectangleXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, x5 As Double, y5 As Double) As Boolean
  TriangleWithinRectangleXY = PointInRectangleXY(X1, Y1, X4, Y4, x5, y5) And _
                                PointInRectangleXY(X2, Y2, X4, Y4, x5, y5) And _
                                PointInRectangleXY(X3, Y3, X4, Y4, x5, y5)
End Function
Public Function TriangleWithinRectangle(Triangle As TTriangle2D, aRectangle As TRectangle) As Boolean
  TriangleWithinRectangle = PointInRectangle(Triangle.P(0), aRectangle) And _
                            PointInRectangle(Triangle.P(1), aRectangle) And _
                            PointInRectangle(Triangle.P(2), aRectangle)
End Function
'(* Endof Triangle Within Rectangle *)

'SegmentWithinRectangle
Public Function SegmentWithinRectangleXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Boolean
  SegmentWithinRectangleXY = PointInRectangleXY(X1, Y1, X3, Y3, X4, Y4) And _
                             PointInRectangleXY(X2, Y2, X3, Y3, X4, Y4)
End Function
Public Function SegmentWithinRectangleS(Segment As TSegment2D, aRectangle As TRectangle) As Boolean
  SegmentWithinRectangleS = SegmentWithinRectangleXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, _
                                                   aRectangle.P(0).X, aRectangle.P(0).Y, aRectangle.P(1).X, aRectangle.P(1).Y)
End Function
'(* Endof Segment Within Rectangle *)

'CircleInCircle
Public Function CircleInCircle(Circle1 As TCircle, Circle2 As TCircle) As Boolean
  CircleInCircle = (PointInCircle2DXYCi(Circle1.X, Circle1.Y, Circle2) And (Circle1.Radius < Circle2.Radius))
End Function
'(* Endof CircleInCircle *)

'IsTangent
Public Function IsTangent(Segment As TSegment2D, aCircle As TCircle) As Boolean
Dim rSqr As Double, drSqr As Double, dSqr As Double, TempSegment As TSegment2D
  TempSegment = Translate2DS(-aCircle.X, -aCircle.Y, Segment)
  rSqr = aCircle.Radius * aCircle.Radius
  drSqr = LayDistance2DS(TempSegment)
  dSqr = Sqr(TempSegment.P(0).X * TempSegment.P(1).Y - TempSegment.P(1).X * TempSegment.P(0).Y)
  IsTangent = IsEqual((rSqr * drSqr - dSqr), 0#)
End Function
'(* Endof IsTangent *)

'PointOfReflection
Public Function PointOfReflection2DXY(Sx1 As Double, Sy1 As Double, Sx2 As Double, Sy2 As Double, P1x As Double, P1y As Double, P2x As Double, P2y As Double, RPx As Double, RPy As Double) As Boolean
Dim ix As Double, iy As Double, P1Px As Double, P1Py As Double, P2Px As Double, P2Py As Double
  PointOfReflection2DXY = False
  If (Not Collinear2D(Sx1, Sy1, Sx2, Sy2, P1x, P1y)) And _
     (Not Collinear2D(Sx1, Sy1, Sx2, Sy2, P2x, P2y)) Then
    Call ClosestPointOnLineFromPointS2DXY(Sx1, Sy1, Sx2, Sy2, P1x, P1y, P1Px, P1Py)
    Call ClosestPointOnLineFromPointS2DXY(Sx1, Sy1, Sx2, Sy2, P2x, P2y, P2Px, P2Py)
    Call Intersect2DXYi(P1x, P1y, P2Px, P2Py, P2x, P2y, P1Px, P1Py, ix, iy)
    Call ClosestPointOnLineFromPointS2DXY(Sx1, Sy1, Sx2, Sy2, ix, iy, RPx, RPy)
    If IsPointCollinear2DXY(Sx1, Sy1, Sx2, Sy2, RPx, RPy) Then PointOfReflection2DXY = True
  End If
End Function
Public Function PointOfReflection2DS(Segment As TSegment2D, P1, P2 As TPoint2D, RP As TPoint2D) As Boolean
  PointOfReflection2DS = PointOfReflection2DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, P1.X, P1.Y, P2.X, P2.Y, RP.X, RP.Y)
End Function
'(* Endof PointOfReflection *)

'Mirror
Public Sub MirrorS2DXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, nX As Double, nY As Double)
  Call ClosestPointOnLineFromPointS2DXY(X1, Y1, X2, Y2, Px, Py, nX, nY)
  nX = Px + 2 * (nX - Px)
  nY = Py + 2 * (nY - Py)
End Sub
Public Function Mirror2DP(Point As TPoint2D, Line As TLine2D) As TPoint2D
  Call MirrorS2DXY(Point.X, Point.Y, Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y, Mirror2DP.X, Mirror2DP.Y)
End Function
Public Function Mirror2DS(Segment As TSegment2D, Line As TLine2D) As TSegment2D
  Mirror2DS.P(0) = Mirror2DP(Segment.P(0), Line)
  Mirror2DS.P(1) = Mirror2DP(Segment.P(1), Line)
End Function
Public Function Mirror2DR(Rectangle As TRectangle, Line As TLine2D) As TRectangle
  Mirror2DR.P(0) = Mirror2DP(Rectangle.P(0), Line)
  Mirror2DR.P(1) = Mirror2DP(Rectangle.P(1), Line)
End Function
Public Function Mirror2DT(Triangle As TTriangle2D, Line As TLine2D) As TTriangle2D
  Mirror2DT.P(0) = Mirror2DP(Triangle.P(0), Line)
  Mirror2DT.P(1) = Mirror2DP(Triangle.P(1), Line)
  Mirror2DT.P(2) = Mirror2DP(Triangle.P(2), Line)
End Function
Public Function Mirror2DQ(Quadix As TQuadix2D, Line As TLine2D) As TQuadix2D
  Mirror2DQ.P(0) = Mirror2DP(Quadix.P(0), Line)
  Mirror2DQ.P(1) = Mirror2DP(Quadix.P(1), Line)
  Mirror2DQ.P(2) = Mirror2DP(Quadix.P(2), Line)
  Mirror2DQ.P(3) = Mirror2DP(Quadix.P(3), Line) 'Achtung hier Fehler in der echten FastGeo!!
End Function
Public Function Mirror2DCi(aCircle As TCircle, Line As TLine2D) As TCircle
  Mirror2DCi.Radius = aCircle.Radius
  Call MirrorS2DXY(aCircle.X, aCircle.Y, Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y, Mirror2DCi.X, Mirror2DCi.Y)
End Function
Public Function MirrorO(Obj As TGeometricObject, Line As TLine2D) As TGeometricObject
  MirrorO.ObjectType = Obj.ObjectType
  Select Case Obj.ObjectType
  Case geoSegment2D:   MirrorO.Segment2D = Mirror2DS(Obj.Segment2D, Line)
  Case geoTriangle2D: MirrorO.Triangle2D = Mirror2DT(Obj.Triangle2D, Line)
  Case geoRectangle:   MirrorO.Rectangle = Mirror2DR(Obj.Rectangle, Line)
  Case geoQuadix2D:     MirrorO.Quadix2D = Mirror2DQ(Obj.Quadix2D, Line)
  Case geoCircle:       MirrorO.aCircle = Mirror2DCi(Obj.aCircle, Line)
  End Select
End Function
'(* Endof Mirror *)

'NonSymmetricMirror
Public Sub NonSymmetricMirror2DXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Ratio As Double, nX As Double, nY As Double)
Dim GeneralRatio As Double
  Call ClosestPointOnLineFromPointS2DXY(X1, Y1, X2, Y2, Px, Py, nX, nY)
  GeneralRatio = 2 * Ratio
  nX = Px + GeneralRatio * (nX - Px)
  nY = Py + GeneralRatio * (nY - Py)
End Sub
Public Function NonSymmetricMirror2DP(Point As TPoint2D, Ratio As Double, Line As TLine2D) As TPoint2D
  Call NonSymmetricMirror2DXY(Point.X, Point.Y, Ratio, Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y, NonSymmetricMirror2DP.X, NonSymmetricMirror2DP.Y)
End Function
Public Function NonSymmetricMirror2DS(Segment As TSegment2D, Ratio As Double, Line As TLine2D) As TSegment2D
  NonSymmetricMirror2DS.P(0) = NonSymmetricMirror2DP(Segment.P(0), Ratio, Line)
  NonSymmetricMirror2DS.P(1) = NonSymmetricMirror2DP(Segment.P(1), Ratio, Line)
End Function
Public Function NonSymmetricMirror2DR(Rectangle As TRectangle, Ratio As Double, Line As TLine2D) As TRectangle
  NonSymmetricMirror2DR.P(0) = NonSymmetricMirror2DP(Rectangle.P(0), Ratio, Line)
  NonSymmetricMirror2DR.P(1) = NonSymmetricMirror2DP(Rectangle.P(1), Ratio, Line)
End Function
Public Function NonSymmetricMirror2DT(Triangle As TTriangle2D, Ratio As Double, Line As TLine2D) As TTriangle2D
  NonSymmetricMirror2DT.P(0) = NonSymmetricMirror2DP(Triangle.P(0), Ratio, Line)
  NonSymmetricMirror2DT.P(1) = NonSymmetricMirror2DP(Triangle.P(1), Ratio, Line)
  NonSymmetricMirror2DT.P(2) = NonSymmetricMirror2DP(Triangle.P(2), Ratio, Line)
End Function
Public Function NonSymmetricMirror2DQ(Quadix As TQuadix2D, Ratio As Double, Line As TLine2D) As TQuadix2D
  NonSymmetricMirror2DQ.P(0) = NonSymmetricMirror2DP(Quadix.P(0), Ratio, Line)
  NonSymmetricMirror2DQ.P(1) = NonSymmetricMirror2DP(Quadix.P(1), Ratio, Line)
  NonSymmetricMirror2DQ.P(2) = NonSymmetricMirror2DP(Quadix.P(2), Ratio, Line)
  NonSymmetricMirror2DQ.P(3) = NonSymmetricMirror2DP(Quadix.P(3), Ratio, Line)
End Function
Public Function NonSymmetricMirror2DCi(aCircle As TCircle, Ratio As Double, Line As TLine2D) As TCircle
  NonSymmetricMirror2DCi.Radius = aCircle.Radius
  Call NonSymmetricMirror2DXY(aCircle.X, aCircle.Y, Ratio, Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y, NonSymmetricMirror2DCi.X, NonSymmetricMirror2DCi.Y)
End Function
Public Function NonSymmetricMirrorO(Obj As TGeometricObject, Ratio As Double, Line As TLine2D) As TGeometricObject
  NonSymmetricMirrorO.ObjectType = Obj.ObjectType
  Select Case Obj.ObjectType
  Case geoSegment2D:  NonSymmetricMirrorO.Segment2D = NonSymmetricMirror2DS(Obj.Segment2D, Ratio, Line)
  Case geoTriangle2D: NonSymmetricMirrorO.Triangle2D = NonSymmetricMirror2DT(Obj.Triangle2D, Ratio, Line)
  Case geoRectangle:  NonSymmetricMirrorO.Rectangle = NonSymmetricMirror2DR(Obj.Rectangle, Ratio, Line)
  Case geoQuadix2D:   NonSymmetricMirrorO.Quadix2D = NonSymmetricMirror2DQ(Obj.Quadix2D, Ratio, Line)
  Case geoCircle:     NonSymmetricMirrorO.aCircle = NonSymmetricMirror2DCi(Obj.aCircle, Ratio, Line)
  End Select
End Function
'(* Endof Non-symmetric Mirror *)

'Distance
Public Function Distance2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
Dim Dx As Double, Dy As Double
  Dx = X2 - X1
  Dy = Y2 - Y1
  Distance2DXY = Sqr(Dx * Dx + Dy * Dy)
End Function
Public Function Distance2DP(Point1 As TPoint2D, Point2 As TPoint2D) As Double
  Distance2DP = Distance2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Function Distance3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As Double
Dim Dx  As Double, Dy  As Double, Dz  As Double
  Dx = X2 - X1
  Dy = Y2 - Y1
  Dz = z2 - z1
  Distance3DXY = Sqr(Dx * Dx + Dy * Dy + Dz * Dz)
End Function
Public Function Distance3DP(Point1 As TPoint3D, Point2 As TPoint3D) As Double
  Distance3DP = Distance3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z)
End Function
Public Function Distance2DPS(Point As TPoint2D, Segment As TSegment2D) As Double
  Distance2DPS = MinimumDistanceFromPointToSegment2D(Point, Segment)
End Function
Public Function Distance2DPR(Point As TPoint2D, aRectangle As TRectangle) As Double
  Distance2DPR = MinimumDistanceFromPointToRectangle(Point, aRectangle)
End Function
Public Function Distance2DPT(Point As TPoint2D, Triangle As TTriangle2D) As Double
  Distance2DPT = MinimumDistanceFromPointToTriangle(Point, Triangle)
End Function
Public Function Distance2DPQ(Point As TPoint2D, Quadix As TQuadix2D) As Double
  Distance2DPQ = VBA.Math.Sqr(LayDistance2DPQ(Point, Quadix))
End Function
Public Function Distance2DLL(Line1 As TLine2D, Line2 As TLine2D) As Double
  Distance2DLL = DistanceLineToLine2D(Line1.P(0).X, Line1.P(0).Y, Line1.P(1).X, Line1.P(1).Y, _
                              Line2.P(0).X, Line2.P(0).Y, Line2.P(1).X, Line2.P(1).Y)
End Function
Public Function Distance3DLL(Line1 As TLine3D, Line2 As TLine3D) As Double
  Distance3DLL = DistanceLineToLine3D(Line1.P(0).X, Line1.P(0).Y, Line1.P(0).z, Line1.P(1).X, Line1.P(1).Y, Line1.P(1).z, _
                              Line2.P(0).X, Line2.P(0).Y, Line2.P(0).z, Line2.P(1).X, Line2.P(1).Y, Line2.P(1).z)
End Function
Public Function Distance2DSS(Segment1 As TSegment2D, Segment2 As TSegment2D) As Double
  Distance2DSS = DistanceSegmentToSegment2D(Segment1.P(0).X, Segment1.P(0).Y, Segment1.P(1).X, Segment1.P(1).Y, _
                                    Segment2.P(0).X, Segment2.P(0).Y, Segment2.P(1).X, Segment2.P(1).Y)
End Function
Public Function Distance3DSS(Segment1 As TSegment3D, Segment2 As TSegment3D) As Double
  Distance3DSS = DistanceSegmentToSegment3D(Segment1.P(0).X, Segment1.P(0).Y, Segment1.P(0).z, Segment1.P(1).X, Segment1.P(1).Y, Segment1.P(1).z, _
                                    Segment2.P(0).X, Segment2.P(0).Y, Segment2.P(0).z, Segment2.P(1).X, Segment2.P(1).Y, Segment2.P(1).z)
End Function
Public Function Distance2DS(Segment As TSegment2D) As Double
  Distance2DS = Distance2DP(Segment.P(0), Segment.P(1))
End Function
Public Function Distance3DS(Segment As TSegment3D) As Double
  Distance3DS = Distance3DP(Segment.P(0), Segment.P(1))
End Function
Public Function Distance2DST(Segment As TSegment2D, Triangle As TTriangle2D) As Double
  Distance2DST = Sqr(LayDistance2DST(Segment, Triangle))
End Function
Public Function Distance3DST(Segment As TSegment3D, Triangle As TTriangle3D) As Double
  Distance3DST = Sqr(LayDistance3DST(Segment, Triangle))
End Function
Public Function Distance2DTT(Triangle1 As TTriangle2D, Triangle2 As TTriangle2D) As Double
Dim i As Long
  Distance2DTT = MinD(MinimumDistanceFromPointToTriangle(Triangle1.P(0), Triangle2), _
                      MinimumDistanceFromPointToTriangle(Triangle2.P(0), Triangle1))
  If IsEqual(Distance2DTT, 0#) Then Exit Function
  Distance2DTT = MinD(Distance2DTT, MinD(MinimumDistanceFromPointToTriangle(Triangle1.P(1), Triangle2), _
                                         MinimumDistanceFromPointToTriangle(Triangle2.P(1), Triangle1)))
  If IsEqual(Distance2DTT, 0#) Then Exit Function
  Distance2DTT = MinD(Distance2DTT, MinD(MinimumDistanceFromPointToTriangle(Triangle1.P(2), Triangle2), _
                                         MinimumDistanceFromPointToTriangle(Triangle2.P(2), Triangle1)))
End Function
Public Function Distance2DTR(Triangle As TTriangle2D, aRectangle As TRectangle) As Double
  If Intersect2DTR(Triangle, aRectangle) Then
    Distance2DTR = 0#
  Else
    Distance2DTR = MinD(MinD(Distance2DST(RectangleEdge(aRectangle, 0), Triangle), _
                             Distance2DST(RectangleEdge(aRectangle, 1), Triangle)), _
                        MinD(Distance2DST(RectangleEdge(aRectangle, 2), Triangle), _
                             Distance2DST(RectangleEdge(aRectangle, 3), Triangle)))
  End If
End Function
Public Function Distance2DRR(Rectangle1 As TRectangle, Rectangle2 As TRectangle) As Double
Dim Rec1 As TRectangle, Rec2 As TRectangle
  If Intersect2DR(Rectangle1, Rectangle2) Then
    Distance2DRR = 0#
  Else
    Rec1 = AABB2DR(Rectangle1)
    Rec2 = AABB2DR(Rectangle2)
    If Rec1.P(2).Y < Rec2.P(1).Y Then
      Distance2DRR = Distance2DSS(EquateSegment2DXY(Rec1.P(0).X, Rec1.P(1).Y, Rec1.P(1).X, Rec1.P(1).Y), EquateSegment2DXY(Rec2.P(0).X, Rec2.P(0).Y, Rec2.P(1).X, Rec2.P(0).Y))
    ElseIf Rec1.P(1).Y > Rec2.P(2).Y Then
      Distance2DRR = Distance2DSS(EquateSegment2DXY(Rec1.P(0).X, Rec1.P(0).Y, Rec1.P(1).X, Rec1.P(0).Y), EquateSegment2DXY(Rec2.P(1).X, Rec2.P(1).Y, Rec2.P(1).X, Rec2.P(1).Y))
    ElseIf Rec1.P(2).X < Rec2.P(1).X Then
      Distance2DRR = Distance2DSS(EquateSegment2DXY(Rec1.P(0).X, Rec1.P(0).Y, Rec1.P(1).X, Rec1.P(1).Y), EquateSegment2DXY(Rec2.P(0).X, Rec2.P(0).Y, Rec2.P(0).X, Rec2.P(1).Y))
    ElseIf Rec1.P(1).X > Rec2.P(2).X Then
      Distance2DRR = Distance2DSS(EquateSegment2DXY(Rec1.P(0).X, Rec1.P(0).Y, Rec1.P(0).X, Rec1.P(1).Y), EquateSegment2DXY(Rec2.P(1).X, Rec2.P(0).Y, Rec2.P(1).X, Rec2.P(1).Y))
    Else
      Distance2DRR = 0#
    End If
  End If
End Function
Public Function Distance2DSR(Segment As TSegment2D, aRectangle As TRectangle) As Double
  Distance2DSR = MinD(MinD(Distance2DSS(Segment, RectangleEdge(aRectangle, 0)), Distance2DSS(Segment, RectangleEdge(aRectangle, 1))), _
                      MinD(Distance2DSS(Segment, RectangleEdge(aRectangle, 2)), Distance2DSS(Segment, RectangleEdge(aRectangle, 3))))
End Function
Public Function Distance2DSCi(Segment As TSegment2D, aCircle As TCircle) As Double
  Distance2DSCi = Distance2DPS(ClosestPointOnCircleFromSegment(aCircle, Segment), Segment)
End Function
Public Function Distance2DTCi(Triangle As TTriangle2D, aCircle As TCircle) As Double
Dim Point1 As TPoint2D, Point2 As TPoint2D
  If Intersect2DTCi(Triangle, aCircle) Then
    Distance2DTCi = 0#
  Else
    Point1 = ClosestPointOnTriangleFromPoint2DTXY(Triangle, aCircle.X, aCircle.Y)
    Point2 = ClosestPointOnCircleFromPoint(aCircle, Point1)
    Distance2DTCi = Distance2DP(Point1, Point2)
  End If
End Function
Public Function Distance2DRCi(aRectangle As TRectangle, aCircle As TCircle) As Double
Dim Point1 As TPoint2D, Point2 As TPoint2D
  If Intersect2DRCi(aRectangle, aCircle) Then
    Distance2DRCi = 0#
  Else
    Point1 = ClosestPointOnRectangleFromPointRXY(aRectangle, aCircle.X, aCircle.Y)
    Point2 = ClosestPointOnCircleFromPoint(aCircle, Point1)
    Distance2DRCi = Distance2DP(Point1, Point2)
  End If
End Function
Public Function Distance2DPCi(Point As TPoint2D, aCircle As TCircle) As Double
  If PointInCircle(Point, aCircle) Then
    Distance2DPCi = 0#
  Else
    Distance2DPCi = Distance2DP(Point, ClosestPointOnCircleFromPoint(aCircle, Point))
  End If
End Function
Public Function Distance2DCi(Circle1 As TCircle, Circle2 As TCircle) As Double
Dim Dist As Double
  Dist = Distance2DXY(Circle1.X, Circle1.Y, Circle2.X, Circle2.Y)
  If Dist > Circle1.Radius + Circle2.Radius Then
    Distance2DCi = Dist - (Circle1.Radius + Circle2.Radius)
  Else
    Distance2DCi = 0#
  End If
End Function
Public Function Distance3DSph(Sphere1 As TSphere, Sphere2 As TSphere) As Double
Dim Dist As Double
  Dist = Distance3DXY(Sphere1.X, Sphere1.Y, Sphere1.z, Sphere2.X, Sphere2.Y, Sphere2.z)
  If Dist > Sphere1.Radius + Sphere2.Radius Then
    Distance3DSph = Dist - (Sphere1.Radius + Sphere2.Radius)
  Else
    Distance3DSph = 0#
  End If
End Function
Public Function DistanceO(Obj1 As TGeometricObject, Obj2 As TGeometricObject) As Double
  DistanceO = 0#
  If (Obj1.ObjectType = geoPoint2D) And (Obj2.ObjectType = geoPoint2D) Then
    DistanceO = Distance2DP(Obj1.Point2D, Obj2.Point2D)
  ElseIf (Obj1.ObjectType = geoPoint3D) And (Obj2.ObjectType = geoPoint3D) Then
    DistanceO = Distance3DP(Obj1.Point3D, Obj2.Point3D)
  ElseIf (Obj1.ObjectType = geoCircle) And (Obj2.ObjectType = geoCircle) Then
    DistanceO = Distance2DCi(Obj1.aCircle, Obj2.aCircle)
  ElseIf (Obj1.ObjectType = geoSegment3D) And (Obj2.ObjectType = geoSegment3D) Then
    DistanceO = Distance3DSS(Obj1.Segment3D, Obj2.Segment3D)
  End If
End Function
'(* End of Distance *)

'LayDistance
Public Function LayDistance2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
Dim Dx As Double, Dy As Double
  Dx = (X2 - X1)
  Dy = (Y2 - Y1)
  LayDistance2DXY = Dx * Dx + Dy * Dy
End Function
Public Function LayDistance2DPP(Point1 As TPoint2D, Point2 As TPoint2D) As Double
  LayDistance2DPP = LayDistance2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Function LayDistance3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As Double
Dim Dx  As Double, Dy  As Double, Dz  As Double
  Dx = X2 - X1
  Dy = Y2 - Y1
  Dz = z2 - z1
  LayDistance3DXY = Dx * Dx + Dy * Dy + Dz * Dz
End Function
Public Function LayDistance3DPP(Point1 As TPoint3D, Point2 As TPoint3D) As Double
  LayDistance3DPP = LayDistance3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z)
End Function
Public Function LayDistance2DPT(Point As TPoint2D, Triangle As TTriangle2D) As Double
  LayDistance2DPT = LayDistance2DPP(Point, ClosestPointOnTriangleFromPoint2D(Triangle, Point))
End Function
Public Function LayDistance2DPQ(Point As TPoint2D, Quadix As TQuadix2D) As Double
  LayDistance2DPQ = LayDistance2DPP(Point, ClosestPointOnQuadixFromPoint2D(Quadix, Point))
End Function
Public Function LayDistance2DSS(Segment1 As TSegment2D, Segment2 As TSegment2D) As Double
  LayDistance2DSS = LayDistanceSegmentToSegment2D(Segment1.P(0).X, Segment1.P(0).Y, Segment1.P(1).X, Segment1.P(1).Y, _
                                                  Segment2.P(0).X, Segment2.P(1).Y, Segment2.P(1).X, Segment2.P(1).Y)
End Function
Public Function LayDistance3DSS(Segment1 As TSegment3D, Segment2 As TSegment3D) As Double
  LayDistance3DSS = LayDistanceSegmentToSegment3D(Segment1.P(0).X, Segment1.P(0).Y, Segment1.P(0).z, Segment1.P(1).X, Segment1.P(1).Y, Segment1.P(1).z, _
                                                  Segment2.P(0).X, Segment2.P(0).Y, Segment2.P(0).z, Segment2.P(1).X, Segment2.P(1).Y, Segment2.P(1).z)
End Function
Public Function LayDistance2DLL(Line1 As TLine2D, Line2 As TLine2D) As Double
  LayDistance2DLL = LayDistanceLineToLine2D(Line1.P(0).X, Line1.P(0).Y, Line1.P(1).X, Line1.P(1).Y, Line2.P(0).X, Line2.P(0).Y, Line2.P(1).X, Line2.P(1).Y)
End Function
Public Function LayDistance3DLL(Line1 As TLine3D, Line2 As TLine3D) As Double
  LayDistance3DLL = LayDistanceLineToLine3D(Line1.P(0).X, Line1.P(0).Y, Line1.P(0).z, Line1.P(1).X, Line1.P(1).Y, Line1.P(1).z, Line2.P(0).X, Line2.P(0).Y, Line2.P(0).z, Line2.P(1).X, Line2.P(1).Y, Line2.P(1).z)
End Function
Public Function LayDistance2DS(Segment As TSegment2D) As Double
  LayDistance2DS = LayDistance2DPP(Segment.P(0), Segment.P(1))
End Function
Public Function LayDistance3DS(Segment As TSegment3D) As Double
  LayDistance3DS = LayDistance3DPP(Segment.P(0), Segment.P(1))
End Function
Public Function LayDistance2DST(Segment As TSegment2D, Triangle As TTriangle2D) As Double
  LayDistance2DST = MinD(MinD(LayDistanceSegmentToSegment2D(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y), _
                              LayDistanceSegmentToSegment2D(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y)), _
                              LayDistanceSegmentToSegment2D(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, Triangle.P(0).X, Triangle.P(0).Y))
End Function
Public Function LayDistance3DST(Segment As TSegment3D, Triangle As TTriangle3D) As Double
  LayDistance3DST = MinD(MinD(LayDistanceSegmentToSegment3D(Segment.P(0).X, Segment.P(0).Y, Segment.P(0).z, Segment.P(1).X, Segment.P(1).Y, Segment.P(1).z, Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(0).z, Triangle.P(0).X, Triangle.P(1).Y, Triangle.P(1).z), _
                              LayDistanceSegmentToSegment3D(Segment.P(0).X, Segment.P(0).Y, Segment.P(0).z, Segment.P(1).X, Segment.P(1).Y, Segment.P(1).z, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(1).z, Triangle.P(2).X, Triangle.P(2).Y, Triangle.P(2).z)), _
                              LayDistanceSegmentToSegment2D(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, Triangle.P(0).X, Triangle.P(0).Y))
End Function
'(* End of LayDistance *)

'ManhattanDistance
Public Function ManhattanDistance2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
  ManhattanDistance2DXY = Abs(X2 - X1) + Abs(Y2 - Y1)
End Function
Public Function ManhattanDistance2DP(Point1 As TPoint2D, Point2 As TPoint2D) As Double
  ManhattanDistance2DP = ManhattanDistance2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Function ManhattanDistance3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As Double
  ManhattanDistance3DXY = Abs(X2 - X1) + Abs(Y2 - Y1) + Abs(z2 - z1)
End Function
Public Function ManhattanDistance3DP(Point1 As TPoint3D, Point2 As TPoint3D) As Double
  ManhattanDistance3DP = ManhattanDistance3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z)
End Function
Public Function ManhattanDistance2DS(Segment As TSegment2D) As Double
  ManhattanDistance2DS = ManhattanDistance2DP(Segment.P(0), Segment.P(1))
End Function
Public Function ManhattanDistance3DS(Segment As TSegment3D) As Double
  ManhattanDistance3DS = ManhattanDistance3DP(Segment.P(0), Segment.P(1))
End Function
Public Function ManhattanDistance2DCi(Circle1 As TCircle, Circle2 As TCircle) As Double
  ManhattanDistance2DCi = ManhattanDistance2DXY(Circle1.X, Circle1.Y, Circle2.X, Circle2.Y)
End Function
'(* Endof ManhattanDistance *)

'VectorSumDistance
Public Function VectorSumDistance2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
  VectorSumDistance2DXY = Abs(X2 - X1) + Abs(Y2 - Y1)
End Function
Public Function VectorSumDistance2DP(Point1 As TPoint2D, Point2 As TPoint2D) As Double
  VectorSumDistance2DP = VectorSumDistance2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Function VectorSumDistance3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As Double
  VectorSumDistance3DXY = Abs(X2 - X1) + Abs(Y2 - Y1) + Abs(z2 - z1)
End Function
Public Function VectorSumDistance3DP(Point1 As TPoint3D, Point2 As TPoint3D) As Double
  VectorSumDistance3DP = VectorSumDistance3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z)
End Function
Public Function VectorSumDistance2DS(Segment As TSegment2D) As Double
  VectorSumDistance2DS = VectorSumDistance2DP(Segment.P(0), Segment.P(1))
End Function
Public Function VectorSumDistance3DS(Segment As TSegment3D) As Double
  VectorSumDistance3DS = VectorSumDistance3DP(Segment.P(0), Segment.P(1))
End Function
Public Function VectorSumDistance2DCi(Circle1 As TCircle, Circle2 As TCircle) As Double
  VectorSumDistance2DCi = VectorSumDistance2DXY(Circle1.X, Circle1.Y, Circle2.X, Circle2.Y)
End Function
Public Function VectorSumDistanceO(Obj1 As TGeometricObject, Obj2 As TGeometricObject) As Double
  VectorSumDistanceO = 0#
  If (Obj1.ObjectType = geoPoint2D) And (Obj2.ObjectType = geoPoint2D) Then
    VectorSumDistanceO = Distance2DP(Obj1.Point2D, Obj2.Point2D)
  ElseIf (Obj1.ObjectType = geoPoint3D) And (Obj2.ObjectType = geoPoint3D) Then
    VectorSumDistanceO = Distance3DP(Obj1.Point3D, Obj2.Point3D)
  ElseIf (Obj1.ObjectType = geoCircle) And (Obj2.ObjectType = geoCircle) Then
    VectorSumDistanceO = Distance2DCi(Obj1.aCircle, Obj2.aCircle)
  End If
End Function
'(* Endof VectorSumDistance *)

'ChebyshevDistance
Public Function ChebyshevDistance2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
  ChebyshevDistance2DXY = MaxD(Abs(X1 - X2), Abs(Y1 - Y2))
End Function
Public Function ChebyshevDistance2DP(Point1 As TPoint2D, Point2 As TPoint2D) As Double
  ChebyshevDistance2DP = ChebyshevDistance2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Function ChebyshevDistance3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As Double
  ChebyshevDistance3DXY = MaxD(MaxD(Abs(X1 - X2), Abs(Y1 - Y2)), Abs(z1 - z2))
End Function
Public Function ChebyshevDistance3DP(Point1 As TPoint3D, Point2 As TPoint3D) As Double
  ChebyshevDistance3DP = ChebyshevDistance3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z)
End Function
Public Function ChebyshevDistance2DS(Segment As TSegment2D) As Double
  ChebyshevDistance2DS = ChebyshevDistance2DP(Segment.P(0), Segment.P(1))
End Function
Public Function ChebyshevDistance3DS(Segment As TSegment3D) As Double
  ChebyshevDistance3DS = ChebyshevDistance3DP(Segment.P(0), Segment.P(1))
End Function
Public Function ChebyshevDistance2DCi(Circle1 As TCircle, Circle2 As TCircle) As Double
  ChebyshevDistance2DCi = MaxD(Abs(Circle1.X - Circle2.X), Abs(Circle1.Y - Circle2.Y))
End Function
'(* Endof ChebyshevDistance *)

'InverseChebyshevDistance
Public Function InverseChebyshevDistance2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
  InverseChebyshevDistance2DXY = MinD(Abs(X1 - X2), Abs(Y1 - Y2))
End Function
Public Function InverseChebyshevDistance2DP(Point1 As TPoint2D, Point2 As TPoint2D) As Double
  InverseChebyshevDistance2DP = InverseChebyshevDistance2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Function InverseChebyshevDistance3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As Double
  InverseChebyshevDistance3DXY = MinD(MinD(Abs(X1 - X2), Abs(Y1 - Y2)), Abs(z1 - z2))
End Function
Public Function InverseChebyshevDistance3DP(Point1 As TPoint3D, Point2 As TPoint3D) As Double
  InverseChebyshevDistance3DP = InverseChebyshevDistance3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z)
End Function
Public Function InverseChebyshevDistance2DS(Segment As TSegment2D) As Double
  InverseChebyshevDistance2DS = InverseChebyshevDistance2DP(Segment.P(0), Segment.P(1))
End Function
Public Function InverseChebyshevDistance3DS(Segment As TSegment3D) As Double
  InverseChebyshevDistance3DS = InverseChebyshevDistance3DP(Segment.P(0), Segment.P(1))
End Function
Public Function InverseChebyshevDistance2DCi(Circle1 As TCircle, Circle2 As TCircle) As Double
  InverseChebyshevDistance2DCi = MinD(Abs(Circle1.X - Circle2.X), Abs(Circle1.Y - Circle2.Y))
End Function
'(* Endof InverseChebyshevDistance *)

'DistanceSegmentToSegment
Public Function DistanceSegmentToSegment2D(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Double
  DistanceSegmentToSegment2D = Sqr(LayDistanceSegmentToSegment2D(X1, Y1, X2, Y2, X3, Y3, X4, Y4))
End Function
Public Function DistanceSegmentToSegment3D(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double) As Double
  DistanceSegmentToSegment3D = Sqr(LayDistanceSegmentToSegment3D(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3, X4, Y4, z4))
End Function
'(* Endof DistanceSegmentToSegment *)

'LayDistanceSegmentToSegment
Public Function LayDistanceSegmentToSegment2D(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Double
Dim ux As Double, uy As Double, Vx As Double, Vy As Double, Wx As Double, Wy As Double
Dim a  As Double, b As Double, c As Double, D As Double, E As Double, dT As Double
Dim sc As Double, sN As Double, sD As Double, tc As Double, tN As Double, tD As Double, Dx As Double, Dy As Double
  ux = X2 - X1
  uy = Y2 - Y1
  
  Vx = X4 - X3
  Vy = Y4 - Y3
  
  Wx = X1 - X3
  Wy = Y1 - Y3
  
  a = (ux * ux + uy * uy)
  b = (ux * Vx + uy * Vy)
  c = (Vx * Vx + Vy * Vy)
  D = (ux * Wx + uy * Wy)
  E = (Vx * Wx + Vy * Wy)
  dT = a * c - b * b
  
  sD = dT
  tD = dT
  
  If IsEqual(dT, 0#) Then
    sN = 0#
    sD = 1#
    tN = E
    tD = c
  Else
    sN = (b * E - c * D)
    tN = (a * E - b * D)
    If sN < 0# Then
      sN = 0#
      tN = E
      tD = c
    ElseIf sN > sD Then
      sN = sD
      tN = E + b
      tD = c
    End If
  End If

  If tN < 0# Then
    tN = 0#
    If -D < 0# Then
      sN = 0#
    ElseIf -D > a Then
      sN = sD
    Else
      sN = -D
      sD = a
    End If
  ElseIf tN > tD Then
    tN = tD
    If (-D + b) < 0# Then
      sN = 0
    ElseIf (-D + b) > a Then
      sN = sD
    Else
      sN = (-D + b)
      sD = a
    End If
  End If
  If IsEqual(sN, 0#) Then
    sc = 0#
  Else
    sc = sN / sD
  End If
  If IsEqual(tN, 0#) Then
    tc = 0#
  Else
    tc = tN / tD
  End If
  Dx = Wx + (sc * ux) - (tc * Vx)
  Dy = Wy + (sc * uy) - (tc * Vy)
  LayDistanceSegmentToSegment2D = Dx * Dx + Dy * Dy
End Function
Public Function LayDistanceSegmentToSegment3D(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double) As Double
Dim ux As Double, uy As Double, uz As Double, Vx As Double, Vy As Double, Vz As Double
Dim Wx As Double, Wy As Double, Wz As Double, a As Double, b As Double, c As Double, D As Double, E As Double
Dim dT As Double, sc As Double, sN As Double, sD As Double, tc As Double, tN As Double, tD As Double, Dx As Double, Dy As Double, Dz As Double
  ux = X2 - X1
  uy = Y2 - Y1
  uz = z2 - z1
  
  Vx = X4 - X3
  Vy = Y4 - Y3
  Vz = z4 - z3
  
  Wx = X1 - X3
  Wy = Y1 - Y3
  Wz = z1 - z3
  
  a = (ux * ux + uy * uy + uz * uz)
  b = (ux * Vx + uy * Vy + uz * Vz)
  c = (Vx * Vx + Vy * Vy + Vz * Vz)
  D = (ux * Wx + uy * Wy + uz * Wz)
  E = (Vx * Wx + Vy * Wy + Vz * Wz)
  dT = a * c - b * b
  
  sD = dT
  tD = dT
  
  If IsEqual(dT, 0#) Then
    sN = 0#
    sD = 1#
    tN = E
    tD = c
  Else
    sN = (b * E - c * D)
    tN = (a * E - b * D)
    If sN < 0# Then
      sN = 0#
      tN = E
      tD = c
    ElseIf sN > sD Then
      sN = sD
      tN = E + b
      tD = c
    End If
  End If
  If tN < 0# Then
    tN = 0#
    If -D < 0# Then
      sN = 0#
    ElseIf -D > a Then
      sN = sD
    Else
      sN = -D
      sD = a
    End If
  ElseIf tN > tD Then
    tN = tD
    If (-D + b) < 0# Then
      sN = 0
    ElseIf (-D + b) > a Then
      sN = sD
    Else
      sN = (-D + b)
      sD = a
    End If
  End If
  If IsEqual(sN, 0#) Then
    sc = 0#
  Else
    sc = sN / sD
  End If
  If IsEqual(tN, 0#) Then
    tc = 0#
  Else
    tc = tN / tD
  End If
  Dx = Wx + (sc * ux) - (tc * Vx)
  Dy = Wy + (sc * uy) - (tc * Vy)
  Dz = Wz + (sc * uz) - (tc * Vz)
  LayDistanceSegmentToSegment3D = Dx * Dx + Dy * Dy + Dz * Dz
End Function
'(* Endof LayDistanceSegmentToSegment *)

'DistanceLineToLine
Public Function DistanceLineToLine2D(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Double
  DistanceLineToLine2D = Sqr(LayDistanceLineToLine2D(X1, Y1, X2, Y2, X3, Y3, X4, Y4))
End Function
Public Function DistanceLineToLine3D(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double) As Double
  DistanceLineToLine3D = Sqr(LayDistanceLineToLine3D(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3, X4, Y4, z4))
End Function
'(* Endof DistanceLineToLine *)

'LayDistanceLineToLine
Public Function LayDistanceLineToLine2D(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Double
Dim ux As Double, uy As Double, Vx As Double, Vy As Double, Wx As Double, Wy As Double
Dim a As Double, b As Double, c As Double, D As Double, E As Double
Dim dT As Double, sc As Double, tc As Double, Dx As Double, Dy As Double
  
  ux = X2 - X1
  uy = Y2 - Y1
  
  Vx = X4 - X3
  Vy = Y4 - Y3
  
  If NotEqual(ux * Vy, uy * Vx) Then
    LayDistanceLineToLine2D = 0#
    Exit Function
  End If
  
  Wx = X1 - X3
  Wy = Y1 - Y3

  a = (ux * ux + uy * uy)
  b = (ux * Vx + uy * Vy)
  c = (Vx * Vx + Vy * Vy)
  D = (ux * Wx + uy * Wy)
  E = (Vx * Wx + Vy * Wy)
  dT = a * c - b * b

  If IsEqual(dT, 0#) Then
    sc = 0#
    If b > c Then
      tc = D / b
    Else
      tc = E / c
    End If
  Else
    sc = (b * E - c * D) / dT
    tc = (a * E - b * D) / dT
  End If
  Dx = Wx + (sc * ux) - (tc * Vx)
  Dy = Wy + (sc * uy) - (tc * Vy)
  LayDistanceLineToLine2D = Dx * Dx + Dy * Dy
End Function
Public Function LayDistanceLineToLine3D(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double) As Double
Dim ux As Double, uy As Double, uz As Double, Vx As Double, Vy As Double, Vz As Double, Wx As Double, Wy As Double, Wz As Double
Dim a As Double, b As Double, c As Double, D As Double, E As Double, dT As Double, sc As Double, tc As Double, Dx As Double, Dy As Double, Dz As Double
  
  ux = X2 - X1
  uy = Y2 - Y1
  uz = z2 - z1
  
  Vx = X4 - X3
  Vy = Y4 - Y3
  Vz = z4 - z3
  
  Wx = X1 - X3
  Wy = Y1 - Y3
  Wz = z1 - z3
  
  a = (ux * ux + uy * uy + uz * uz)
  b = (ux * Vx + uy * Vy + uz * Vz)
  c = (Vx * Vx + Vy * Vy + Vz * Vz)
  D = (ux * Wx + uy * Wy + uz * Wz)
  E = (Vx * Wx + Vy * Wy + Vz * Wz)
  dT = a * c - b * b
  
  If IsEqual(dT, 0#) Then
    sc = 0#
    If b > c Then
      tc = D / b
    Else
      tc = E / c
    End If
  Else
    sc = (b * E - c * D) / dT
    tc = (a * E - b * D) / dT
  End If
  Dx = Wx + (sc * ux) - (tc * Vx)
  Dy = Wy + (sc * uy) - (tc * Vy)
  Dz = Wz + (sc * uz) - (tc * Vz)
  LayDistanceLineToLine3D = Dx * Dx + Dy * Dy + Dz * Dz
End Function
'(* Endof LayDistanceLineToLine *)

'TriangleType
Public Function TriangleType2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As eTriangletype
  If IsEquilateralTriangle2DXY(X1, Y1, X2, Y2, X3, Y3) Then
    TriangleType2DXY = etEquilateral
  Else
    If IsIsoscelesTriangle2DXY(X1, Y1, X2, Y2, X3, Y3) Then
      TriangleType2DXY = etIsosceles
    Else
      If IsRightTriangle2DXY(X1, Y1, X2, Y2, X3, Y3) Then
        TriangleType2DXY = etRight
      Else
        If IsScaleneTriangle2DXY(X1, Y1, X2, Y2, X3, Y3) Then
          TriangleType2DXY = etScalene
        Else
          If IsObtuseTriangle2DXY(X1, Y1, X2, Y2, X3, Y3) Then
            TriangleType2DXY = etObtuse
          Else
            TriangleType2DXY = etUnknown
          End If
        End If
      End If
    End If
  End If
End Function
Public Function TriangleType3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As eTriangletype
  If IsEquilateralTriangle3DXY(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3) Then
    TriangleType3DXY = etEquilateral
  Else
    If IsIsoscelesTriangle3DXY(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3) Then
      TriangleType3DXY = etIsosceles
    Else
      If IsRightTriangle3DXY(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3) Then
        TriangleType3DXY = etRight
      Else
        If IsScaleneTriangle3DXY(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3) Then
          TriangleType3DXY = etScalene
        Else
          If IsObtuseTriangle3DXY(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3) Then
            TriangleType3DXY = etObtuse
          Else
            TriangleType3DXY = etUnknown
          End If
        End If
      End If
    End If
  End If
End Function
Public Function TriangleType2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As eTriangletype
  TriangleType2DP = TriangleType2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function TriangleType3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As eTriangletype
  TriangleType3DP = TriangleType3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z)
End Function
Public Function TriangleType2DT(Triangle As TTriangle2D) As eTriangletype 'OK!!
  TriangleType2DT = TriangleType2DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
Public Function TriangleType3DT(Triangle As TTriangle3D) As eTriangletype 'OK!!
  TriangleType3DT = TriangleType3DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
'(* Endof Triangletype *)


'IsEquilateralTriangle
'ist ein gleichseitiges Dreieck
Public Function IsEquilateralTriangle2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Boolean
Dim d1 As Double, d2 As Double, d3 As Double
  d1 = LayDistance2DXY(X1, Y1, X2, Y2)
  d2 = LayDistance2DXY(X2, Y2, X3, Y3)
  d3 = LayDistance2DXY(X3, Y3, X1, Y1)
  IsEquilateralTriangle2DXY = (IsEqual(d1, d2) And IsEqual(d2, d3))
End Function
Public Function IsEquilateralTriangle3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As Boolean
Dim d1 As Double, d2 As Double, d3 As Double
  d1 = LayDistance3DXY(X1, Y1, z1, X2, Y2, z2)
  d2 = LayDistance3DXY(X2, Y2, z2, X3, Y3, z3)
  d3 = LayDistance3DXY(X3, Y3, z3, X1, Y1, z1)
  IsEquilateralTriangle3DXY = (IsEqual(d1, d2) And IsEqual(d2, d3))
End Function
Public Function IsEquilateralTriangle2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Boolean
  IsEquilateralTriangle2DP = IsEquilateralTriangle2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function IsEquilateralTriangle3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As Boolean
  IsEquilateralTriangle3DP = IsEquilateralTriangle3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z)
End Function
Public Function IsEquilateralTriangle2DT(Triangle As TTriangle2D) As Boolean 'OK!!
  IsEquilateralTriangle2DT = IsEquilateralTriangle2DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
Public Function IsEquilateralTriangle3DT(Triangle As TTriangle3D) As Boolean 'OK!!
  IsEquilateralTriangle3DT = IsEquilateralTriangle3DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
'(* Endof IsEquilateralTriangle *)

'IsIsoscelesTriangle
Public Function IsIsoscelesTriangle2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Boolean
Dim d1 As Double, d2 As Double, d3 As Double
  d1 = LayDistance2DXY(X1, Y1, X2, Y2)
  d2 = LayDistance2DXY(X2, Y2, X3, Y3)
  d3 = LayDistance2DXY(X3, Y3, X1, Y1)
  IsIsoscelesTriangle2DXY = ((IsEqual(d1, d2) Or IsEqual(d1, d3)) And NotEqual(d2, d3)) Or (IsEqual(d2, d3) And NotEqual(d2, d1))
End Function
Public Function IsIsoscelesTriangle3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As Boolean
Dim d1 As Double, d2 As Double, d3 As Double
  d1 = LayDistance3DXY(X1, Y1, z1, X2, Y2, z2)
  d2 = LayDistance3DXY(X2, Y2, z2, X3, Y3, z3)
  d3 = LayDistance3DXY(X3, Y3, z3, X1, Y1, z1)
  IsIsoscelesTriangle3DXY = ((IsEqual(d1, d2) Or IsEqual(d1, d3)) And NotEqual(d2, d3)) Or (IsEqual(d2, d3) And NotEqual(d2, d1))
End Function
Public Function IsIsoscelesTriangle2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Boolean
  IsIsoscelesTriangle2DP = IsIsoscelesTriangle2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function IsIsoscelesTriangle3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As Boolean
  IsIsoscelesTriangle3DP = IsIsoscelesTriangle3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z)
End Function
Public Function IsIsoscelesTriangle2D(Triangle As TTriangle2D) As Boolean
  IsIsoscelesTriangle2D = IsIsoscelesTriangle2DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
Public Function IsIsoscelesTriangle3D(Triangle As TTriangle3D) As Boolean
  IsIsoscelesTriangle3D = IsIsoscelesTriangle3DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
'(* Endof IsIsoscelesTriangle *)

'IsRightTriangle
Public Function IsRightTriangle2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Boolean
Dim d1 As Double, d2 As Double, d3 As Double
  d1 = LayDistance2DXY(X1, Y1, X2, Y2)
  d2 = LayDistance2DXY(X2, Y2, X3, Y3)
  d3 = LayDistance2DXY(X3, Y3, X1, Y1)
  IsRightTriangle2DXY = (IsEqual(d1 + d2, d3) Or _
                         IsEqual(d1 + d3, d2) Or _
                         IsEqual(d3 + d2, d1))
End Function
Public Function IsRightTriangle3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As Boolean
Dim d1 As Double, d2 As Double, d3 As Double
  d1 = LayDistance3DXY(X1, Y1, z1, X2, Y2, z2)
  d2 = LayDistance3DXY(X2, Y2, z2, X3, Y3, z3)
  d3 = LayDistance3DXY(X3, Y3, z3, X1, Y1, z1)
  IsRightTriangle3DXY = (IsEqual(d1 + d2, d3) Or _
                         IsEqual(d1 + d3, d2) Or _
                         IsEqual(d3 + d2, d1))
End Function
Public Function IsRightTriangle2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Boolean
  IsRightTriangle2DP = IsRightTriangle2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function IsRightTriangle3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As Boolean
  IsRightTriangle3DP = IsRightTriangle3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z)
End Function
Public Function IsRightTriangle2D(Triangle As TTriangle2D) As Boolean
  IsRightTriangle2D = IsRightTriangle2DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
Public Function IsRightTriangle3D(Triangle As TTriangle3D) As Boolean
  IsRightTriangle3D = IsRightTriangle3DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
'(* Endof IsRightTriangle *)

'IsScaleneTriangle
Public Function IsScaleneTriangle2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Boolean
Dim d1 As Double, d2 As Double, d3 As Double
  d1 = LayDistance2DXY(X1, Y1, X2, Y2)
  d2 = LayDistance2DXY(X2, Y2, X3, Y3)
  d3 = LayDistance2DXY(X3, Y3, X1, Y1)
  IsScaleneTriangle2DXY = NotEqual(d1, d2) And NotEqual(d2, d3) And NotEqual(d3, d1)
End Function
Public Function IsScaleneTriangle3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As Boolean
Dim d1 As Double, d2 As Double, d3 As Double
  d1 = LayDistance3DXY(X1, Y1, z1, X2, Y2, z2)
  d2 = LayDistance3DXY(X2, Y2, z2, X3, Y3, z3)
  d3 = LayDistance3DXY(X3, Y3, z3, X1, Y1, z1)
  IsScaleneTriangle3DXY = NotEqual(d1, d2) And NotEqual(d2, d3) And NotEqual(d3, d1)
End Function
Public Function IsScaleneTriangle2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Boolean
  IsScaleneTriangle2DP = IsScaleneTriangle2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function IsScaleneTriangle3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As Boolean
  IsScaleneTriangle3DP = IsScaleneTriangle3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z)
End Function
Public Function IsScaleneTriangle2D(Triangle As TTriangle2D) As Boolean
  IsScaleneTriangle2D = IsScaleneTriangle2DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
Public Function IsScaleneTriangle3D(Triangle As TTriangle3D) As Boolean
  IsScaleneTriangle3D = IsScaleneTriangle3DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
'(* Endof IsScaleneTriangle *)

'IsObtuseTriangle
Public Function IsObtuseTriangle2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Boolean
Dim a1 As Double, a2 As Double, a3 As Double
  a1 = VertexAngle2DXY(X1, Y1, X2, Y2, X3, Y3)
  a2 = VertexAngle2DXY(X3, Y3, X1, Y1, X2, Y2)
  a3 = VertexAngle2DXY(X2, Y2, X3, Y3, X1, Y1)
  IsObtuseTriangle2DXY = (a1 > 90#) Or (a2 > 90#) Or (a3 > 90#)
End Function
Public Function IsObtuseTriangle3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As Boolean
Dim a1 As Double, a2 As Double, a3 As Double
  a1 = VertexAngle3DXY(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3)
  a2 = VertexAngle3DXY(X3, Y3, z3, X1, Y1, z1, X2, Y2, z2)
  a3 = VertexAngle3DXY(X2, Y2, z2, X3, Y3, z3, X1, Y1, z1)
  IsObtuseTriangle3DXY = (a1 > 90#) Or (a2 > 90#) Or (a3 > 90#)
End Function
Public Function IsObtuseTriangle2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Boolean
  IsObtuseTriangle2DP = IsObtuseTriangle2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function IsObtuseTriangle3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As Boolean
  IsObtuseTriangle3DP = IsObtuseTriangle3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z)
End Function
Public Function IsObtuseTriangle2D(Triangle As TTriangle2D) As Boolean
  IsObtuseTriangle2D = IsObtuseTriangle2DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
Public Function IsObtuseTriangle3D(Triangle As TTriangle3D) As Boolean
  IsObtuseTriangle3D = IsObtuseTriangle3DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
'(* Endof IsObtuseTriangle *)

'TriangleEdge
Public Function TriangleEdge2D(Triangle As TTriangle2D, Edge As Long) As TSegment2D
  Select Case Edge
  Case 1: TriangleEdge2D = EquateSegment2DP(Triangle.P(0), Triangle.P(1))
  Case 2: TriangleEdge2D = EquateSegment2DP(Triangle.P(1), Triangle.P(2))
  Case 3: TriangleEdge2D = EquateSegment2DP(Triangle.P(2), Triangle.P(0))
  End Select
End Function
Public Function TriangleEdge3D(Triangle As TTriangle3D, Edge As Long) As TSegment3D
  Select Case Edge
  Case 1: TriangleEdge3D = EquateSegment3DP(Triangle.P(0), Triangle.P(1))
  Case 2: TriangleEdge3D = EquateSegment3DP(Triangle.P(1), Triangle.P(2))
  Case 3: TriangleEdge3D = EquateSegment3DP(Triangle.P(2), Triangle.P(0))
  End Select
End Function
'(* Triangle Edge *)

'RectangleEdge
Public Function RectangleEdge(Rectangle As TRectangle, Edge As Long) As TSegment2D
  Select Case Edge
  Case 1: RectangleEdge = EquateSegment2DXY(Rectangle.P(0).X, Rectangle.P(0).Y, Rectangle.P(0).X, Rectangle.P(0).Y)
  Case 2: RectangleEdge = EquateSegment2DXY(Rectangle.P(1).X, Rectangle.P(0).Y, Rectangle.P(1).X, Rectangle.P(1).Y)
  Case 3: RectangleEdge = EquateSegment2DXY(Rectangle.P(1).X, Rectangle.P(1).Y, Rectangle.P(0).X, Rectangle.P(1).Y)
  Case 4: RectangleEdge = EquateSegment2DXY(Rectangle.P(0).X, Rectangle.P(1).Y, Rectangle.P(0).X, Rectangle.P(0).Y)
  End Select
End Function
'(* Rectangle Edge*)

'PointInTriangle
Public Function PointInTriangle2DXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Boolean
Dim Or1 As Long, Or2 As Long, Or3 As Long
  Or1 = Orientation2D(X1, Y1, X2, Y2, Px, Py)
  Or2 = Orientation2D(X2, Y2, X3, Y3, Px, Py)
  If (Or1 * Or2) = -1 Then
    PointInTriangle2DXY = False
  Else
    Or3 = Orientation2D(X3, Y3, X1, Y1, Px, Py)
    If (Or1 = Or3) Or (Or3 = 0) Then
      PointInTriangle2DXY = True
    ElseIf Or1 = 0 Then
      PointInTriangle2DXY = (Or2 * Or3) >= 0
    ElseIf Or2 = 0 Then
      PointInTriangle2DXY = (Or1 * Or3) >= 0
    Else
      PointInTriangle2DXY = False
    End If
  End If
  '(*
  '  Note As  The following code es the same as above,
  '       but is less efficient time-wise.
  '  Or1 = Orientation(x1, y1, x2, y2, Px, Py)
  '  Or2 = Orientation(x2, y2, x3, y3, Px, Py)
  '  Or3 = Orientation(x3, y3, x1, y1, Px, Py)
  '
  '  If (Or1 = Or2) And (Or2 = Or3) Then
  '   Result = True
  '  ElseIf Or1 = 0 Then
  '   Result = (Or2 * Or3) >= 0
  '  ElseIf Or2 = 0 Then
  '   Result = (Or1 * Or3) >= 0
  '  ElseIf Or3 = 0 Then
  '   Result = (Or1 * Or2) >= 0
  '  Else
  '   Result = False
  '*)
End Function
Public Function PointInTriangle2DXYT(X As Double, Y As Double, Triangle As TTriangle2D) As Boolean
  PointInTriangle2DXYT = PointInTriangle2DXY(X, Y, Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y)
End Function
Public Function PointInTriangle(Point As TPoint2D, Triangle As TTriangle2D) As Boolean
  PointInTriangle = PointInTriangle2DXY(Point.X, Point.Y, Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y)
End Function
'(* Endof PointInTriangle *)

'PointInCircle
Public Function PointInCircle2DXYR(Px As Double, Py As Double, cX As Double, cY As Double, Radius As Double) As Boolean
  PointInCircle2DXYR = (LayDistance2DXY(Px, Py, cX, cY) <= (Radius * Radius))
End Function
Public Function PointInCircle2DXYCi(Px As Double, Py As Double, aCircle As TCircle) As Boolean
  PointInCircle2DXYCi = PointInCircle2DXYR(Px, Py, aCircle.X, aCircle.Y, aCircle.Radius)
End Function
Public Function PointInCircle(Point As TPoint2D, aCircle As TCircle) As Boolean
  PointInCircle = PointInCircle2DXYCi(Point.X, Point.Y, aCircle)
End Function
'(* End of PointInCircle *)

'PointOnCircle:
Public Function PointOnCircle2DXYCi(Px As Double, Py As Double, aCircle As TCircle) As Boolean
  PointOnCircle2DXYCi = IsEqual(LayDistance2DXY(Px, Py, aCircle.X, aCircle.Y), (aCircle.Radius * aCircle.Radius))
End Function
Public Function PointOnCircle(Point As TPoint2D, aCircle As TCircle) As Boolean
  PointOnCircle = PointOnCircle2DXYCi(Point.X, Point.Y, aCircle)
End Function
'(* End of PointOnCircle *)

'TriangleInCircle
Public Function TriangleInCircle(Triangle As TTriangle2D, aCircle As TCircle) As Boolean
  TriangleInCircle = PointInCircle(Triangle.P(0), aCircle) And _
                     PointInCircle(Triangle.P(1), aCircle) And _
                     PointInCircle(Triangle.P(2), aCircle)
End Function
'(* Endof TriangleInCircle *)

'TriangleOutsideCircle
Public Function TriangleOutsideCircle(Triangle As TTriangle2D, aCircle As TCircle) As Boolean
  TriangleOutsideCircle = (Not PointInCircle(Triangle.P(0), aCircle)) And _
                          (Not PointInCircle(Triangle.P(1), aCircle)) And _
                          (Not PointInCircle(Triangle.P(2), aCircle))
End Function
'(* Endof TriangleOutsideCircle *)

'TriangleEncompassesCircle
Public Function TriangleEncompassesCircle(Triangle As TTriangle2D, aCircle As TCircle) As Boolean
  TriangleEncompassesCircle = TriangleOutsideCircle(Triangle, aCircle) And _
            PointInCircle(ClosestPointOnTriangleFromPoint2DTXY(Triangle, aCircle.X, aCircle.Y), aCircle)
End Function
'(* Endof TriangleEncompassesCircle *)

'RectangleInCircle
Public Function RectangleInCircle(Rectangle As TRectangle, aCircle As TCircle) As Boolean
  RectangleInCircle = PointInCircle2DXYCi(Rectangle.P(0).X, Rectangle.P(0).Y, aCircle) And _
                      PointInCircle2DXYCi(Rectangle.P(1).X, Rectangle.P(1).Y, aCircle) And _
                      PointInCircle2DXYCi(Rectangle.P(0).X, Rectangle.P(1).Y, aCircle) And _
                      PointInCircle2DXYCi(Rectangle.P(1).X, Rectangle.P(0).Y, aCircle)
End Function
'(* Endof RectangleInCircle *)

'RectangleOutsideCircle
Public Function RectangleOutsideCircle(Rectangle As TRectangle, aCircle As TCircle) As Boolean
  RectangleOutsideCircle = (Not PointInCircle2DXYCi(Rectangle.P(0).X, Rectangle.P(0).Y, aCircle)) And _
                           (Not PointInCircle2DXYCi(Rectangle.P(1).X, Rectangle.P(1).Y, aCircle)) And _
                           (Not PointInCircle2DXYCi(Rectangle.P(0).X, Rectangle.P(1).Y, aCircle)) And _
                           (Not PointInCircle2DXYCi(Rectangle.P(1).X, Rectangle.P(0).Y, aCircle))
End Function
'(* Endof RectangleInCircle *)

'QuadixInCircle
Public Function QuadixInCircle(Quadix As TQuadix2D, aCircle As TCircle) As Boolean
  QuadixInCircle = PointInCircle(Quadix.P(0), aCircle) And _
                   PointInCircle(Quadix.P(1), aCircle) And _
                   PointInCircle(Quadix.P(2), aCircle) And _
                   PointInCircle(Quadix.P(3), aCircle)
End Function
'(* Endof QuadixInCircle *)

'QuadixOutsideCircle
Public Function QuadixOutsideCircle(Quadix As TQuadix2D, aCircle As TCircle) As Boolean
  QuadixOutsideCircle = (Not PointInCircle(Quadix.P(0), aCircle)) And _
                        (Not PointInCircle(Quadix.P(1), aCircle)) And _
                        (Not PointInCircle(Quadix.P(2), aCircle)) And _
                        (Not PointInCircle(Quadix.P(3), aCircle))
End Function
'(* Endof QuadixInCircle *)

'PointInThreePointCircle
Public Function PointInThreePointCircle2DXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Boolean
Dim a11 As Double, a12 As Double, a21 As Double, a22 As Double
Dim Dx1 As Double, Dx2 As Double, Dx3 As Double
Dim Dy3 As Double, Dy1 As Double, Dy2 As Double
  Dx1 = X1 - Px
  Dx2 = X2 - Px
  Dx3 = X3 - Px
  Dy1 = Y2 - Py
  Dy2 = Y3 - Py
  Dy3 = Y1 - Py
  a11 = Dx3 * Dy1 - Dx2 * Dy2
  a12 = Dx3 * Dy3 - Dx1 * Dy2
  a21 = Dx2 * (X2 - X3) + Dy1 * (Y2 - Y3)
  a22 = Dx1 * (X1 - X3) + Dy3 * (Y1 - Y3)
  PointInThreePointCircle2DXY = ((a11 * a22 - a21 * a12) <= 0)
End Function
Public Function PointInThreePointCircle2DP(Point As TPoint2D, Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Boolean
  PointInThreePointCircle2DP = PointInThreePointCircle2DXY(Point.X, Point.Y, Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Function
Public Function PointInThreePointCircle2DPT(Point As TPoint2D, Triangle As TTriangle2D) As Boolean
  PointInThreePointCircle2DPT = PointInThreePointCircle2DP(Point, Triangle.P(1), Triangle.P(2), Triangle.P(3))
End Function
'(* Endof Point In Three Point Circle *)

'PointInRectangle
Public Function PointInRectangleXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Boolean
  PointInRectangleXY = ((X1 <= Px) And (Px <= X2) And (Y1 <= Py) And (Py <= Y2)) Or _
            ((X2 <= Px) And (Px <= X1) And (Y2 <= Py) And (Py <= Y1))
End Function
Public Function PointInRectanglePXY(Point As TPoint2D, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Boolean
  PointInRectanglePXY = PointInRectangleXY(Point.X, Point.Y, X1, Y1, X2, Y2)
End Function
Public Function PointInRectangleXYR(Px As Double, Py As Double, aRectangle As TRectangle) As Boolean
  PointInRectangleXYR = PointInRectangleXY(Px, Py, aRectangle.P(0).X, aRectangle.P(0).Y, aRectangle.P(1).X, aRectangle.P(1).Y)
End Function
Public Function PointInRectangle(Point As TPoint2D, aRectangle As TRectangle) As Boolean
  PointInRectangle = PointInRectangleXY(Point.X, Point.Y, aRectangle.P(0).X, aRectangle.P(0).Y, aRectangle.P(1).X, aRectangle.P(1).Y)
End Function
'(* Endof PointInRectangle *)

'TriangleInRectangle
Public Function TriangleInRectangle(Triangle As TTriangle2D, aRectangle As TRectangle) As Boolean
  TriangleInRectangle = PointInRectangle(Triangle.P(0), aRectangle) And _
                        PointInRectangle(Triangle.P(1), aRectangle) And _
                        PointInRectangle(Triangle.P(2), aRectangle)
End Function
'(* Endof TriangleInRectangle *)

'TriangleOutsideRectangle
Public Function TriangleOutsideRectangle(Triangle As TTriangle2D, aRectangle As TRectangle) As Boolean
  TriangleOutsideRectangle = (Not PointInRectangle(Triangle.P(0), aRectangle)) And _
                             (Not PointInRectangle(Triangle.P(1), aRectangle)) And _
                             (Not PointInRectangle(Triangle.P(2), aRectangle))
End Function
'(* Endof TriangleOutsideRectangle *)

'QuadixInRectangle
Public Function QuadixInRectangle(Quadix As TQuadix2D, aRectangle As TRectangle) As Boolean
  QuadixInRectangle = PointInRectangle(Quadix.P(0), aRectangle) And _
                      PointInRectangle(Quadix.P(1), aRectangle) And _
                      PointInRectangle(Quadix.P(2), aRectangle) And _
                      PointInRectangle(Quadix.P(3), aRectangle)
End Function
'(* Endof QuadixInRectangle *)

'QuadixOutsideRectangle
Public Function QuadixOutsideRectangle(Quadix As TQuadix2D, aRectangle As TRectangle) As Boolean
  QuadixOutsideRectangle = (Not PointInRectangle(Quadix.P(0), aRectangle)) And _
                           (Not PointInRectangle(Quadix.P(1), aRectangle)) And _
                           (Not PointInRectangle(Quadix.P(2), aRectangle)) And _
                           (Not PointInRectangle(Quadix.P(3), aRectangle))
End Function
'(* Endof QuadixOutsideRectangle *)

'PointInQuadix
Public Function PointInQuadix2DXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Boolean
Dim Or1 As Long, Or2 As Long, Or3 As Long, Or4 As Long
  Or1 = Orientation2D(X1, Y1, X2, Y2, Px, Py)
  Or2 = Orientation2D(X2, Y2, X3, Y3, Px, Py)
  Or3 = Orientation2D(X3, Y3, X4, Y4, Px, Py)
  Or4 = Orientation2D(X4, Y4, X1, Y1, Px, Py)
  If (Or1 = Or2) And (Or2 = Or3) And (Or3 = Or4) Then
    PointInQuadix2DXY = True
  ElseIf Or1 = 0 Then
    PointInQuadix2DXY = (Or2 * Or4) = 0
  ElseIf Or2 = 0 Then
    PointInQuadix2DXY = (Or1 * Or3) = 0
  ElseIf Or3 = 0 Then
    PointInQuadix2DXY = (Or2 * Or4) = 0
  ElseIf Or4 = 0 Then
    PointInQuadix2DXY = (Or1 * Or3) = 0
  Else
    PointInQuadix2DXY = False
  End If
End Function
Public Function PointInQuadix2DP(Point As TPoint2D, Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D) As Boolean
  PointInQuadix2DP = PointInQuadix2DXY(Point.X, Point.Y, Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Point4.X, Point4.Y)
End Function
Public Function PointInQuadix2DXYQ(X As Double, Y As Double, Quadix As TQuadix2D) As Boolean
  PointInQuadix2DXYQ = PointInQuadix2DXY(X, Y, Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(2).X, Quadix.P(2).Y, Quadix.P(3).X, Quadix.P(3).Y)
End Function
Public Function PointInQuadix(Point As TPoint2D, Quadix As TQuadix2D) As Boolean
  PointInQuadix = PointInQuadix2DP(Point, Quadix.P(0), Quadix.P(1), Quadix.P(2), Quadix.P(3))
End Function
'(* Endof PointInQuadix *)

'TriangleInQuadix
Public Function TriangleInQuadix(Triangle As TTriangle2D, Quadix As TQuadix2D) As Boolean
  TriangleInQuadix = PointInQuadix(Triangle.P(0), Quadix) And _
                     PointInQuadix(Triangle.P(1), Quadix) And _
                     PointInQuadix(Triangle.P(2), Quadix)
End Function
'(* Endof TriangleInQuadix *)

'TriangleOutsideQuadix
Public Function TriangleOutsideQuadix(Triangle As TTriangle2D, Quadix As TQuadix2D) As Boolean
  TriangleOutsideQuadix = (Not PointInQuadix(Triangle.P(0), Quadix)) And _
                          (Not PointInQuadix(Triangle.P(1), Quadix)) And _
                          (Not PointInQuadix(Triangle.P(2), Quadix))
End Function
'(* Endof TriangleOutsideQuadix *)

'PointInSphere
Public Function PointInSphere()
  '
End Function

Public Function PointOnSphere(Point3D As TPoint3D, Sphere As TSphere) As Boolean
  PointOnSphere = IsEqual(LayDistance3DXY(Point3D.X, Point3D.Y, Point3D.z, Sphere.z, Sphere.Y, Sphere.z), (Sphere.Radius * Sphere.Radius))
End Function
'(* Endof PointOnSphere *)

'PolyhedronInSphere
Public Function PolyhedronInSphere(Polygon As TPolyhedron, Sphere As TSphere) As TInclusion
Dim i As Long, j As Long, Count As Long, RealCount As Long
  RealCount = 0
  Count = 0
  For i = 0 To UBound(Polygon.Arr)
    Call Inc(RealCount, UBound(Polygon.Arr(i).Arr))
    For j = 0 To UBound(Polygon.Arr(i).Arr)
      If PointInSphere(Polygon.Arr(i).Arr(j), Sphere) Then
        Call Inc(Count)
      End If
    Next
  Next
  PolyhedronInSphere = ePartially
  If Count = 0 Then
    PolyhedronInSphere = eOutside
  ElseIf Count = RealCount Then
    PolyhedronInSphere = eFully
  End If
End Function
'(* Endof PolyhedronInSphere *)

'PointOnPerimeter
Public Function PointOnPerimeter2DXYS(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Boolean
  '//Result = (((Px = x1) or (Px = x2)) and ((Py = y1) or (Py = y2)))
  PointOnPerimeter2DXYS = (((IsEqual(Px, X1) Or IsEqual(Px, X2)) And ((Py >= Y1) And (Py <= Y2))) Or _
                           ((IsEqual(Py, Y1) Or IsEqual(Py, Y2)) And ((Px >= X1) And (Px <= X2))))
End Function
Public Function PointOnPerimeter2DXYT(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Optional Robust As Boolean = False) As Boolean
  PointOnPerimeter2DXYT = (IsPointCollinear2DXY(X1, Y1, X2, Y2, Px, Py, Robust) Or _
                           IsPointCollinear2DXY(X2, Y2, X3, Y3, Px, Py, Robust) Or _
                           IsPointCollinear2DXY(X3, Y3, X1, Y1, Px, Py, Robust))
End Function
Public Function PointOnPerimeter2DXYQ(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As Boolean
  PointOnPerimeter2DXYQ = (IsPointCollinear2DXY(X1, Y1, X2, Y2, Px, Py) Or _
                           IsPointCollinear2DXY(X2, Y2, X3, Y3, Px, Py) Or _
                           IsPointCollinear2DXY(X3, Y3, X4, Y4, Px, Py) Or _
                           IsPointCollinear2DXY(X4, Y4, X1, Y1, Px, Py))
End Function
Public Function PointOnPerimeter2DR(Point As TPoint2D, aRectangle As TRectangle) As Boolean
  PointOnPerimeter2DR = PointOnPerimeter2DXYS(Point.X, Point.Y, aRectangle.P(0).X, aRectangle.P(0).Y, aRectangle.P(1).X, aRectangle.P(1).Y)
End Function
Public Function PointOnPerimeter2DT(Point As TPoint2D, Triangle As TTriangle2D) As Boolean
  PointOnPerimeter2DT = PointOnPerimeter2DXYT(Point.X, Point.Y, Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y)
End Function
Public Function PointOnPerimeter2DQ(Point As TPoint2D, Quadix As TQuadix2D) As Boolean
  PointOnPerimeter2DQ = PointOnPerimeter2DXYQ(Point.X, Point.Y, Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(2).X, Quadix.P(2).Y, Quadix.P(3).X, Quadix.P(3).Y)
End Function
Public Function PointOnPerimeter2DCi(Point As TPoint2D, aCircle As TCircle) As Boolean
  '//Result = (LayDistance(Point.x,Point.y,Circle.x,Circle.y) = (Circle.Radius * Circle.Radius))
  PointOnPerimeter2DCi = IsEqual(LayDistance2DXY(Point.X, Point.Y, aCircle.X, aCircle.Y), (aCircle.Radius * aCircle.Radius))
End Function
Public Function PointOnPerimeter3DSph(Point As TPoint3D, Sphere As TSphere) As Boolean
  '//Result = (LayDistance(Point.x,Point.y,Point.z,Sphr.x,Sphr.y,Sphr.z) = (Sphr.Radius * Sphr.Radius))
  PointOnPerimeter3DSph = IsEqual(LayDistance3DXY(Point.X, Point.Y, Point.z, Sphere.X, Sphere.Y, Sphere.z), (Sphere.Radius * Sphere.Radius))
End Function
Public Function PointOnPerimeter2DPg(Point As TPoint2D, Polygon As TPolygon2D) As Boolean
  PointOnPerimeter2DPg = PointOnPolygonXY(Point.X, Point.Y, Polygon)
End Function
Public Function PointOnPerimeter2DO(Point As TPoint2D, Obj As TGeometricObject) As Boolean
  Select Case Obj.ObjectType
  Case geoRectangle:  PointOnPerimeter2DO = PointOnPerimeter2DR(Point, Obj.Rectangle)
  Case geoTriangle2D: PointOnPerimeter2DO = PointOnPerimeter2DT(Point, Obj.Triangle2D)
  Case geoQuadix2D:   PointOnPerimeter2DO = PointOnPerimeter2DQ(Point, Obj.Quadix2D)
  Case geoCircle:     PointOnPerimeter2DO = PointOnPerimeter2DCi(Point, Obj.aCircle)
  Case Else:          PointOnPerimeter2DO = False
  End Select
End Function
Public Function PointOnPerimeter3DO(Point As TPoint3D, Obj As TGeometricObject) As Boolean
  Select Case Obj.ObjectType
  Case geoSphere: PointOnPerimeter3DO = PointOnPerimeter3DSph(Point, Obj.Sphere)
  Case Else
    PointOnPerimeter3DO = False
  End Select
End Function
'(* Endof PointOnPerimeter *)

'PointInObject (ist bis jetzt nur 2D!!)
Public Function PointInObject2DS(Point As TPoint2D, Segment As TSegment2D) As Boolean
  PointInObject2DS = IsPointCollinear2DSP(Segment, Point)
End Function
Public Function PointInObject2DL(Point As TPoint2D, Line As TLine2D) As Boolean
  PointInObject2DL = Collinear2DP(Line.P(0), Line.P(1), Point)
End Function
Public Function PointInObject2DR(Point As TPoint2D, aRectangle As TRectangle) As Boolean
  PointInObject2DR = PointInRectangle(Point, aRectangle)
End Function
Public Function PointInObject2DT(Point As TPoint2D, Triangle As TTriangle2D) As Boolean
  PointInObject2DT = PointInTriangle(Point, Triangle)
End Function
Public Function PointInObject2DQ(Point As TPoint2D, Quadix As TQuadix2D) As Boolean
  PointInObject2DQ = PointInQuadix(Point, Quadix)
End Function
Public Function PointInObject2DCi(Point As TPoint2D, aCircle As TCircle) As Boolean
  PointInObject2DCi = PointInCircle(Point, aCircle)
End Function
Public Function PointInObject2DPg(Point As TPoint2D, Polygon As TPolygon2D) As Boolean
  PointInObject2DPg = PointInPolygon(Point, Polygon)
End Function
Public Function PointInObject2DO(Point As TPoint2D, Obj As TGeometricObject) As Boolean
  Select Case Obj.ObjectType
  Case geoSegment2D:  PointInObject2DO = PointInObject2DS(Point, Obj.Segment2D)
  Case geoLine2D:     PointInObject2DO = PointInObject2DL(Point, Obj.Line2D)
  Case geoRectangle:  PointInObject2DO = PointInObject2DR(Point, Obj.Rectangle)
  Case geoTriangle2D: PointInObject2DO = PointInObject2DT(Point, Obj.Triangle2D)
  Case geoQuadix2D:   PointInObject2DO = PointInObject2DQ(Point, Obj.Quadix2D)
  Case geoCircle:     PointInObject2DO = PointInObject2DCi(Point, Obj.aCircle)
  Case geoPolygon2D:  PointInObject2DO = PointInObject2DPg(Point, Obj.Polygon2D)
  Case Else:          PointInObject2DO = False
  End Select
End Function
'(* Endof Point In Object *)

'GeometricSpan 'Achtung hier -1!! vordefiniert
Public Function GeometricSpan2D(Point() As TPoint2D) As Double
Dim TempDistance As Double, i As Long, j As Long
  GeometricSpan2D = -1
  For i = 0 To UBound(Point) - 1 '- 2
    For j = (i + 1) To UBound(Point)
      TempDistance = LayDistance2DPP(Point(i), Point(j))
      If TempDistance > GeometricSpan2D Then
        GeometricSpan2D = TempDistance
      End If
    Next
  Next
  GeometricSpan2D = Sqr(GeometricSpan2D)
End Function
Public Function GeometricSpan3D(Point() As TPoint3D) As Double
Dim TempDistance As Double, i As Long, j As Long
  GeometricSpan3D = -1
  If UBound(Point) < 2 Then Exit Function
  For i = 0 To UBound(Point) - 1 '- 2
    For j = (i + 1) To UBound(Point)
      TempDistance = LayDistance3DPP(Point(i), Point(j))
      If TempDistance > GeometricSpan3D Then
        GeometricSpan3D = TempDistance
      End If
    Next
  Next
  GeometricSpan3D = Sqr(GeometricSpan3D)
End Function
'(* Endof 3D Geometric Span *)

'CreateEquilateralTriangle
Public Sub CreateEquilateralTriangleS2DXY(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByRef X3 As Double, ByRef Y3 As Double)
Dim Sin60 As Double: Sin60 = 0.866025403784439  '0.86602540378443864676372317075294
Dim Cos60 As Double: Cos60 = 0.5                '0.50000000000000000000000000000000
  '(* Translate for x1, y1 to be origin *)
  X2 = X2 - X1
  Y2 = Y2 - Y1
  '(* Rotate 60 degrees and translate back *)
  X3 = ((X2 * Cos60) - (Y2 * Sin60)) + X1
  Y3 = ((Y2 * Cos60) + (X2 * Sin60)) + Y1
End Sub
Public Sub CreateEquilateralTriangleS2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D)
  Call CreateEquilateralTriangleS2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y)
End Sub
Public Function CreateEquilateralTriangle2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As TTriangle2D
  CreateEquilateralTriangle2DXY.P(0).X = X1
  CreateEquilateralTriangle2DXY.P(0).Y = Y1
  CreateEquilateralTriangle2DXY.P(1).X = X2
  CreateEquilateralTriangle2DXY.P(1).Y = Y2
  Call CreateEquilateralTriangleS2DXY(X1, Y1, X2, Y2, CreateEquilateralTriangle2DXY.P(2).X, CreateEquilateralTriangle2DXY.P(2).Y)
End Function
Public Function CreateEquilateralTriangle2DP(Point1 As TPoint2D, Point2 As TPoint2D) As TTriangle2D
  CreateEquilateralTriangle2DP.P(0) = Point1
  CreateEquilateralTriangle2DP.P(1) = Point2
  Call CreateEquilateralTriangleS2DP(CreateEquilateralTriangle2DP.P(0), CreateEquilateralTriangle2DP.P(1), CreateEquilateralTriangle2DP.P(2))
End Function
Public Function CreateEquilateralTriangle2DXYSl(cX As Double, cY As Double, SideLength As Double) As TTriangle2D
  CreateEquilateralTriangle2DXYSl = CenterAtLocation2DTXY(CreateEquilateralTriangle2DXY(-SideLength * 0.5, 0#, SideLength * 0.5, 0#), cX, cY)
End Function
Public Function CreateEquilateralTriangle2DCSl(CenterPoint As TPoint2D, SideLength As Double) As TTriangle2D
  CreateEquilateralTriangle2DCSl = CreateEquilateralTriangle2DXYSl(CenterPoint.X, CenterPoint.Y, SideLength)
End Function
'(* Endof Create Equilateral Triangle *)

'TorricelliPoint
Public Sub TorricelliPoint2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Px As Double, Py As Double)
Dim OETx1 As Double, OETy1 As Double, OETx2 As Double, OETy2 As Double
  '(*
  '  Proven by Cavalieri in this book "Exercitationes geometricae" 1647.
  '  The theory goes, If the triangle has an angle of 120 degrees or more
  '  the toricelli point lies at the vertex of the large angle. Otherwise
  '  the point a which the Simpson lines intersect is said to be the optimal
  '  solution.
  '  to find an intersection in 2D, all that is needed is 2 lines (segments),
  '  hence not all three of the Simpson lines are calculated.
  '*)
  '//If VertexAngle(x1,y1,x2,y2,x3,y3) >= 120.0 then
  If GreaterThanOrEqual(VertexAngle2DXY(X1, Y1, X2, Y2, X3, Y3), 120#) Then
    Px = X2
    Py = Y2
    Exit Sub
  '//ElseIf VertexAngle(x3,y3,x1,y1,x2,y2) >= 120.0 then
  ElseIf GreaterThanOrEqual(VertexAngle2DXY(X3, Y3, X1, Y1, X2, Y2), 120#) Then
    Px = X1
    Py = Y1
    Exit Sub
  '//ElseIf VertexAngle(x2,y2,x3,y3,x1,y1) >= 120.0 then
  ElseIf GreaterThanOrEqual(VertexAngle2DXY(X2, Y2, X3, Y3, X1, Y1), 120#) Then
    Px = X3
    Py = Y3
    Exit Sub
  Else
    If Orientation2D(X1, Y1, X2, Y2, X3, Y3) = RightHandSide Then
      Call CreateEquilateralTriangleS2DXY(X1, Y1, X2, Y2, OETx1, OETy1)
      Call CreateEquilateralTriangleS2DXY(X2, Y2, X3, Y3, OETx2, OETy2)
    Else
      Call CreateEquilateralTriangleS2DXY(X2, Y2, X1, Y1, OETx1, OETy1)
      Call CreateEquilateralTriangleS2DXY(X3, Y3, X2, Y2, OETx2, OETy2)
    End If
    Call IntersectionPoint2DXYN(OETx1, OETy1, X3, Y3, OETx2, OETy2, X1, Y1, Px, Py)
  End If
End Sub
Public Function TorricelliPoint2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As TPoint2D
  Call TorricelliPoint2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, TorricelliPoint2DP.X, TorricelliPoint2DP.Y)
End Function
Public Function TorricelliPoint2DT(Triangle As TTriangle2D) As TPoint2D
  TorricelliPoint2DT = TorricelliPoint2DP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
'(* Endof Create Torricelli Point *)

'Incenter
Public Sub IncenterS2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Px As Double, Py As Double)
Dim Perim  As Double, Side12 As Double, Side23 As Double, Side31 As Double
  Side12 = Distance2DXY(X1, Y1, X2, Y2)
  Side23 = Distance2DXY(X2, Y2, X3, Y3)
  Side31 = Distance2DXY(X3, Y3, X1, Y1)
  '(* Using Heron's S=UR *)
  Perim = 1# / (Side12 + Side23 + Side31)
  Px = (Side23 * X1 + Side31 * X2 + Side12 * X3) * Perim
  Py = (Side23 * Y1 + Side31 * Y2 + Side12 * Y3) * Perim
End Sub
Public Sub IncenterS2DT(Triangle As TTriangle2D, Px As Double, Py As Double)
  Call IncenterS2DXY(Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, Px, Py)
End Sub
Public Function Incenter2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As TPoint2D
  Call IncenterS2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, Incenter2DP.X, Incenter2DP.Y)
End Function
Public Function Incenter2DT(Triangle As TTriangle2D) As TPoint2D
  Call IncenterS2DXY(Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, Incenter2DT.X, Incenter2DT.Y)
End Function
'(* Endof Incenter *)

'Circumcenter
Public Sub CircumcenterXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Px As Double, Py As Double)
Dim a As Double, b As Double, c As Double, D As Double, E As Double, F As Double, G As Double
  a = X2 - X1
  b = Y2 - Y1
  c = X3 - X1
  D = Y3 - Y1
  E = a * (X1 + X2) + b * (Y1 + Y2)
  F = c * (X1 + X3) + D * (Y1 + Y3)
  G = 2# * (a * (Y3 - Y2) - b * (X3 - X2))
  If IsEqual(G, 0#) Then Exit Sub
  Px = (D * E - b * F) / G
  Py = (a * F - c * E) / G
End Sub
Public Function CircumcenterP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As TPoint2D
  Call CircumcenterXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Point3.X, Point3.Y, CircumcenterP.X, CircumcenterP.Y)
End Function
Public Function CircumcenterT(Triangle As TTriangle2D) As TPoint2D
  Call CircumcenterXY(Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, CircumcenterT.X, CircumcenterT.Y)
End Function
'(* Endof Circumcenter *)


'Circumcircle
Public Function CircumcircleP(P1 As TPoint2D, P2 As TPoint2D, P3 As TPoint2D) As TCircle
  Call CircumcenterXY(P1.X, P1.Y, P2.X, P2.Y, P3.X, P3.Y, CircumcircleP.X, CircumcircleP.Y)
  CircumcircleP.Radius = Distance2DXY(P1.X, P1.Y, CircumcircleP.X, CircumcircleP.Y)
End Function
Public Function CircumcircleT(Triangle As TTriangle2D) As TCircle
  CircumcircleT = CircumcircleP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
'(* Endof TriangleCircumCircle *)

'InscribedCircle
Public Function InscribedCircleXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As TCircle
Dim Perimeter As Double
Dim Side12 As Double, Side23 As Double, Side31 As Double
  Side12 = Distance2DXY(X1, Y1, X2, Y2)
  Side23 = Distance2DXY(X2, Y2, X3, Y3)
  Side31 = Distance2DXY(X3, Y3, X1, Y1)
  '(* Using Heron's S = UR *)
  Perimeter = 1# / (Side12 + Side23 + Side31)
  With InscribedCircleXY
    .X = (Side23 * X1 + Side31 * X2 + Side12 * X3) * Perimeter
    .Y = (Side23 * Y1 + Side31 * Y2 + Side12 * Y3) * Perimeter
    .Radius = 0.5 * Sqr((-Side12 + Side23 + Side31) * (Side12 - Side23 + Side31) * (Side12 + Side23 - Side31) * Perimeter)
  End With
End Function
Public Function InscribedCircleP(P1 As TPoint2D, P2 As TPoint2D, P3 As TPoint2D) As TCircle
  InscribedCircleP = InscribedCircleXY(P1.X, P1.Y, P2.X, P2.Y, P3.X, P3.Y)
End Function
Public Function InscribedCircleT(Triangle As TTriangle2D) As TCircle
  InscribedCircleT = InscribedCircleP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
'(* Endof InscribedCircle *)

'ClosestPointOnSegmentFromPoint:
Public Sub ClosestPointOnSegmentFromPointS2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double, nX As Double, nY As Double)
Dim Vx As Double, Vy As Double, Wx As Double, Wy As Double, c1 As Double, c2 As Double, Ratio As Double
  Vx = X2 - X1
  Vy = Y2 - Y1
  Wx = Px - X1
  Wy = Py - Y1
  c1 = Vx * Wx + Vy * Wy
  If c1 <= 0 Then
    nX = X1
    nY = Y1
    Exit Sub
  End If
  c2 = Vx * Vx + Vy * Vy
  If c2 <= c1 Then
    nX = X2
    nY = Y2
    Exit Sub
  End If
  Ratio = c1 / c2
  nX = X1 + Ratio * Vx
  nY = Y1 + Ratio * Vy
End Sub
Public Sub ClosestPointOnSegmentFromPointS3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, Px As Double, Py As Double, Pz As Double, nX As Double, nY As Double, Nz As Double)
Dim Vx As Double, Vy As Double, Vz As Double, Wx As Double, Wy As Double, Wz As Double, c1 As Double, c2 As Double, Ratio As Double
  Vx = X2 - X1
  Vy = Y2 - Y1
  Vz = z2 - z1
  Wx = Px - X1
  Wy = Py - Y1
  Wz = Pz - z1
  c1 = Vx * Wx + Vy * Wy + Vz * Wz
  If c1 <= 0 Then
    nX = X1
    nY = Y1
    Nz = z1
    Exit Sub
  End If
  c2 = Vx * Vx + Vy * Vy + Vz * Vz
  If c2 <= c1 Then
    nX = X2
    nY = Y2
    Nz = z2
    Exit Sub
  End If
  Ratio = c1 / c2
  nX = X1 + Ratio * Vx
  nY = Y1 + Ratio * Vy
  Nz = z1 + Ratio * Vz
End Sub
Public Function ClosestPointOnSegmentFromPoint2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double) As TPoint2D
  Call ClosestPointOnSegmentFromPointS2DXY(X1, Y1, X2, Y2, Px, Py, ClosestPointOnSegmentFromPoint2DXY.X, ClosestPointOnSegmentFromPoint2DXY.Y)
End Function
Public Function ClosestPointOnSegmentFromPoint3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, Px As Double, Py As Double, Pz As Double) As TPoint3D
  Call ClosestPointOnSegmentFromPointS3DXY(X1, Y1, z1, X2, Y2, z2, Px, Py, Pz, ClosestPointOnSegmentFromPoint3DXY.X, ClosestPointOnSegmentFromPoint3DXY.Y, ClosestPointOnSegmentFromPoint3DXY.z)
End Function
Public Function ClosestPointOnSegmentFromPoint2DSP(Segment As TSegment2D, Point As TPoint2D) As TPoint2D
  Call ClosestPointOnSegmentFromPointS2DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Point.X, Point.Y, ClosestPointOnSegmentFromPoint2DSP.X, ClosestPointOnSegmentFromPoint2DSP.Y)
End Function
Public Function ClosestPointOnSegmentFromPoint3DSP(Segment As TSegment3D, Point As TPoint3D) As TPoint3D
  Call ClosestPointOnSegmentFromPointS3DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(0).z, Segment.P(1).X, Segment.P(1).Y, Segment.P(1).z, Point.X, Point.Y, Point.z, ClosestPointOnSegmentFromPoint3DSP.X, ClosestPointOnSegmentFromPoint3DSP.Y, ClosestPointOnSegmentFromPoint3DSP.z)
End Function
'(* End of ClosestPointOnSegmentFromPoint *)

'ClosestPointOnLineFromPoint
Public Sub ClosestPointOnLineFromPointS2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double, nX As Double, nY As Double)
Dim Vx As Double, Vy As Double, Wx As Double, Wy As Double, c1 As Double, c2 As Double, Ratio As Double
  Vx = X2 - X1
  Vy = Y2 - Y1
  Wx = Px - X1
  Wy = Py - Y1
  c1 = Vx * Wx + Vy * Wy
  c2 = Vx * Vx + Vy * Vy
  Ratio = c1 / c2
  nX = X1 + Ratio * Vx
  nY = Y1 + Ratio * Vy
End Sub
Public Sub ClosestPointOnLineFromPointS3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, Px As Double, Py As Double, Pz As Double, nX As Double, nY As Double, Nz As Double)
Dim Vx As Double, Vy As Double, Vz As Double, Wx As Double, Wy As Double, Wz As Double, c1 As Double, c2 As Double, Ratio As Double
  Vx = X2 - X1
  Vy = Y2 - Y1
  Vz = z2 - z1
  Wx = Px - X1
  Wy = Py - Y1
  Wz = Pz - z1
  c1 = Vx * Wx + Vy * Wy + Vz * Wz
  c2 = Vx * Vx + Vy * Vy + Vz * Vz
  Ratio = c1 / c2
  nX = X1 + Ratio * Vx
  nY = Y1 + Ratio * Vy
  Nz = z1 + Ratio * Vz
End Sub
Public Function ClosestPointOnLineFromPoint2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double) As TPoint2D
  Call ClosestPointOnLineFromPointS2DXY(X1, Y1, X2, Y2, Px, Py, ClosestPointOnLineFromPoint2DXY.X, ClosestPointOnLineFromPoint2DXY.Y)
End Function
Public Function ClosestPointOnLineFromPoint3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, Px As Double, Py As Double, Pz As Double) As TPoint3D
  Call ClosestPointOnLineFromPointS3DXY(X1, Y1, z1, X2, Y2, z2, Px, Py, Pz, ClosestPointOnLineFromPoint3DXY.X, ClosestPointOnLineFromPoint3DXY.Y, ClosestPointOnLineFromPoint3DXY.z)
End Function
Public Function ClosestPointOnLineFromPoint2DLP(Line As TLine2D, Point As TPoint2D) As TPoint2D
  Call ClosestPointOnLineFromPointS2DXY(Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y, Point.X, Point.Y, ClosestPointOnLineFromPoint2DLP.X, ClosestPointOnLineFromPoint2DLP.Y)
End Function
Public Function ClosestPointOnLineFromPoint3DLP(Line As TLine3D, Point As TPoint3D) As TPoint3D
  Call ClosestPointOnLineFromPointS3DXY(Line.P(0).X, Line.P(0).Y, Line.P(0).z, Line.P(1).X, Line.P(1).Y, Line.P(1).z, Point.X, Point.Y, Point.z, ClosestPointOnLineFromPoint3DLP.X, ClosestPointOnLineFromPoint3DLP.Y, ClosestPointOnLineFromPoint3DLP.z)
End Function
'(* End of ClosestPointOnLineFromPoint *)

'ClosestPointOnTriangleFromPoint
Public Sub ClosestPointOnTriangleFromPointS2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Px As Double, Py As Double, nX As Double, nY As Double)
  If Orientation2D(X1, Y1, X2, Y2, Px, Py) <> Orientation2D(X1, Y1, X2, Y2, X3, Y3) Then
    Call ClosestPointOnSegmentFromPointS2DXY(X1, Y1, X2, Y2, Px, Py, nX, nY)
    Exit Sub
  End If
  If Orientation2D(X2, Y2, X3, Y3, Px, Py) <> Orientation2D(X2, Y2, X3, Y3, X1, Y1) Then
    Call ClosestPointOnSegmentFromPointS2DXY(X2, Y2, X3, Y3, Px, Py, nX, nY)
    Exit Sub
  End If
  If Orientation2D(X3, Y3, X1, Y1, Px, Py) <> Orientation2D(X3, Y3, X1, Y1, X2, Y2) Then
    Call ClosestPointOnSegmentFromPointS2DXY(X3, Y3, X1, Y1, Px, Py, nX, nY)
    Exit Sub
  End If
  nX = Px
  nY = Py
End Sub
Public Function ClosestPointOnTriangleFromPoint2DTXY(Triangle As TTriangle2D, Px As Double, Py As Double) As TPoint2D
  Call ClosestPointOnTriangleFromPointS2DXY(Triangle.P(0).X, Triangle.P(0).Y, _
                                            Triangle.P(1).X, Triangle.P(1).Y, _
                                            Triangle.P(2).X, Triangle.P(2).Y, _
                                            Px, Py, _
                                            ClosestPointOnTriangleFromPoint2DTXY.X, ClosestPointOnTriangleFromPoint2DTXY.Y)
End Function
Public Function ClosestPointOnTriangleFromPoint2D(Triangle As TTriangle2D, Point As TPoint2D) As TPoint2D
  ClosestPointOnTriangleFromPoint2D = ClosestPointOnTriangleFromPoint2DTXY(Triangle, Point.X, Point.Y)
End Function
Public Function ClosestPointOnTriangleFromPoint3D(Triangle As TTriangle3D, Point As TPoint3D) As TPoint3D
Dim Leydist1 As Double, Leydist2 As Double, Leydist3 As Double
  Leydist1 = LayDistance3DPP(Triangle.P(0), Point)
  Leydist2 = LayDistance3DPP(Triangle.P(1), Point)
  Leydist3 = LayDistance3DPP(Triangle.P(2), Point)
  '(* Edge 3 is the longest *)
  If GreaterThanOrEqual(Leydist3, Leydist1) And GreaterThanOrEqual(Leydist3, Leydist2) Then
    Call ClosestPointOnSegmentFromPointS2DXY(Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y, Point.X, Point.Y, ClosestPointOnTriangleFromPoint3D.X, ClosestPointOnTriangleFromPoint3D.Y)
    Exit Function
  End If '???müßte hier nicht            S3DXY stehen?
  '(* Edge 2 is the longest *)
  If GreaterThanOrEqual(Leydist2, Leydist1) And GreaterThanOrEqual(Leydist2, Leydist3) Then
    Call ClosestPointOnSegmentFromPointS2DXY(Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(2).X, Triangle.P(2).Y, Point.X, Point.Y, ClosestPointOnTriangleFromPoint3D.X, ClosestPointOnTriangleFromPoint3D.Y)
    Exit Function
  End If
  '(* Edge 1 is the longest *)
  If GreaterThanOrEqual(Leydist1, Leydist2) And GreaterThanOrEqual(Leydist1, Leydist3) Then
    Call ClosestPointOnSegmentFromPointS2DXY(Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, Point.X, Point.Y, ClosestPointOnTriangleFromPoint3D.X, ClosestPointOnTriangleFromPoint3D.Y)
    Exit Function
  End If
End Function
'(* Closest Point On Triangle 3D From Point *)

Public Sub ClosestPointOnRectangleFromPointXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Px As Double, Py As Double, nX As Double, nY As Double)
  If (Px < MinD(X1, X2)) Then
    nX = MinD(X1, X2)
  ElseIf (Px > MaxD(X1, X2)) Then
    nX = MaxD(X1, X2)
  Else
    nX = Px
  End If
  If (Py < MinD(Y1, Y2)) Then
    nY = MinD(Y1, Y2)
  ElseIf (Py > MaxD(Y1, Y2)) Then
    nY = MaxD(Y1, Y2)
  Else
    nY = Py
  End If
End Sub
Public Function ClosestPointOnRectangleFromPointRXY(Rectangle As TRectangle, Px As Double, Py As Double) As TPoint2D
  Call ClosestPointOnRectangleFromPointXY(Rectangle.P(0).X, Rectangle.P(0).Y, Rectangle.P(1).X, Rectangle.P(1).Y, Px, Py, ClosestPointOnRectangleFromPointRXY.X, ClosestPointOnRectangleFromPointRXY.Y)
End Function
Public Function ClosestPointOnRectangleFromPointRP(Rectangle As TRectangle, Point As TPoint2D) As TPoint2D
  ClosestPointOnRectangleFromPointRP = ClosestPointOnRectangleFromPointRXY(Rectangle, Point.X, Point.Y)
End Function
'(* Closest Point On Rectangle From Point *)

Public Function ClosestPointOnQuadixFromPoint2D(Quadix As TQuadix2D, Point As TPoint2D) As TPoint2D
Dim MinDist As Double, TempDist As Double, TempPoint As TPoint2D
  If PointInQuadix(Point, Quadix) Then
    ClosestPointOnQuadixFromPoint2D = Point
    Exit Function
  End If
  Call ClosestPointOnSegmentFromPointS2DXY(Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(1).X, Quadix.P(1).Y, Point.X, Point.Y, ClosestPointOnQuadixFromPoint2D.X, ClosestPointOnQuadixFromPoint2D.Y)
  MinDist = Distance2DP(ClosestPointOnQuadixFromPoint2D, Point)
  Call ClosestPointOnSegmentFromPointS2DXY(Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(2).X, Quadix.P(2).Y, Point.X, Point.Y, TempPoint.X, TempPoint.Y)
  TempDist = Distance2DP(TempPoint, Point)
  If MinDist > TempDist Then
    MinDist = TempDist
    ClosestPointOnQuadixFromPoint2D = TempPoint
  End If
  Call ClosestPointOnSegmentFromPointS2DXY(Quadix.P(2).X, Quadix.P(2).Y, Quadix.P(3).X, Quadix.P(3).Y, Point.X, Point.Y, TempPoint.X, TempPoint.Y)
  TempDist = Distance2DP(TempPoint, Point)
  If MinDist > TempDist Then
    MinDist = TempDist
    ClosestPointOnQuadixFromPoint2D = TempPoint
  End If
  Call ClosestPointOnSegmentFromPointS2DXY(Quadix.P(3).X, Quadix.P(3).Y, Quadix.P(0).X, Quadix.P(0).Y, Point.X, Point.Y, TempPoint.X, TempPoint.Y)
  TempDist = Distance2DP(TempPoint, Point)
  If MinDist > TempDist Then
    ClosestPointOnQuadixFromPoint2D = TempPoint
  End If
End Function
'(* Closest Point On Quadix 2D From Point *)

'ClosestPointOnQuadixFromPoint
Public Function ClosestPointOnQuadixFromPoint3D(Quadix As TQuadix3D, Point As TPoint3D) As TPoint3D
'(*
'Dim Leydist1 As Double, Leydist2 As Double, Leydist3 As Double, Leydist4 As Double
'*)

 '(*
 ' Leydist1 = LayDistance(Quadix.P(0), Point)
 ' Leydist2 = LayDistance(Quadix.P(1), Point)
 ' Leydist3 = LayDistance(Quadix.P(2), Point)
 ' Leydist4 = LayDistance(Quadix.P(3), Point)
 '*)
End Function
'(* Closest Point On Quadix 3D From Point *)

'ClosestPointOnCircleFromPoint
Public Function ClosestPointOnCircleFromPoint(aCircle As TCircle, Point As TPoint2D) As TPoint2D
Dim Ratio As Double, Dx As Double, Dy As Double
  Dx = Point.X - aCircle.X
  Dy = Point.Y - aCircle.Y
  Ratio = aCircle.Radius / Sqr(Dx * Dx + Dy * Dy)
  ClosestPointOnCircleFromPoint.X = aCircle.X + Ratio * Dx
  ClosestPointOnCircleFromPoint.Y = aCircle.Y + Ratio * Dy
End Function
'(* Closest Point On Circle From Point *)

'ClosestPointOnCircleFromSegment
Public Function ClosestPointOnCircleFromSegment(aCircle As TCircle, Segment As TSegment2D) As TPoint2D
Dim nX As Double, nY As Double, Ratio As Double
  Call ClosestPointOnSegmentFromPointS2DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, aCircle.X, aCircle.Y, nX, nY)
  Ratio = aCircle.Radius / Distance2DXY(aCircle.X, aCircle.Y, nX, nY)
  ClosestPointOnCircleFromSegment.X = aCircle.X + Ratio * (nX - aCircle.X)
  ClosestPointOnCircleFromSegment.Y = aCircle.Y + Ratio * (nY - aCircle.Y)
End Function
'(* Endof ClosestPointOnCircle *)

'ClosestPointOnSphereFromPoint
Public Function ClosestPointOnSphereFromPoint(Sphere As TSphere, Point As TPoint3D) As TPoint3D
Dim Ratio As Double, Dx As Double, Dy As Double, Dz As Double
  Dx = Point.X - Sphere.X
  Dy = Point.Y - Sphere.Y
  Dz = Point.z - Sphere.z
  Ratio = Sphere.Radius / Sqr(Dx * Dx + Dy * Dy + Dz * Dz)
  ClosestPointOnSphereFromPoint.X = Sphere.X + Ratio * Dx
  ClosestPointOnSphereFromPoint.Y = Sphere.Y + Ratio * Dy
  ClosestPointOnSphereFromPoint.z = Sphere.z + Ratio * Dz
End Function
'(* Closest Point On Sphere From Point *)

'ClosestPointOnSphereFromSegment
Public Function ClosestPointOnSphereFromSegment(Sphere As TSphere, Segment As TSegment3D) As TPoint3D
Dim nX As Double, nY As Double, Nz As Double, Ratio As Double
  Call ClosestPointOnSegmentFromPointS3DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(0).z, Segment.P(1).X, Segment.P(1).Y, Segment.P(1).z, Sphere.X, Sphere.Y, Sphere.z, nX, nY, Nz)
  Ratio = Sphere.Radius / Distance3DXY(Sphere.X, Sphere.Y, Sphere.z, nX, nY, Nz)
  ClosestPointOnSphereFromSegment.X = Sphere.X + Ratio * (nX - Sphere.X)
  ClosestPointOnSphereFromSegment.Y = Sphere.Y + Ratio * (nY - Sphere.Y)
  ClosestPointOnSphereFromSegment.z = Sphere.z + Ratio * (Nz - Sphere.z)
End Function
'(* Endof ClosestPointOnCircle *)

'ClosestPointOnAABBFromPoint
Public Function ClosestPointOnAABBFromPoint(Rectangle As TRectangle, Point As TPoint2D) As TPoint2D
  ClosestPointOnAABBFromPoint = Point
  If Point.X <= Rectangle.P(0).X Then
    ClosestPointOnAABBFromPoint.X = Rectangle.P(0).X
  ElseIf Point.X >= Rectangle.P(1).X Then
    ClosestPointOnAABBFromPoint.X = Rectangle.P(1).X
  End If
  If Point.Y <= Rectangle.P(0).Y Then
    ClosestPointOnAABBFromPoint.Y = Rectangle.P(0).Y
  ElseIf Point.Y >= Rectangle.P(1).Y Then
    ClosestPointOnAABBFromPoint.Y = Rectangle.P(1).Y
  End If
End Function
'(* Closest Point On AABB From Point *)

'MinimumDistanceFromPointToSegment
Public Function MinimumDistanceFromPointToSegment2DXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
Dim nX As Double, nY As Double
  Call ClosestPointOnSegmentFromPointS2DXY(X1, Y1, X2, Y2, Px, Py, nX, nY)
  MinimumDistanceFromPointToSegment2DXY = Distance2DXY(Px, Py, nX, nY)
End Function
Public Function MinimumDistanceFromPointToSegment2D(Point As TPoint2D, Segment As TSegment2D) As Double
  MinimumDistanceFromPointToSegment2D = MinimumDistanceFromPointToSegment2DXY(Point.X, Point.Y, Segment.P(1).X, Segment.P(1).Y, Segment.P(2).X, Segment.P(2).Y)
End Function
Public Function MinimumDistanceFromPointToSegment3DXY(Px As Double, Py As Double, Pz As Double, X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As Double
Dim nX As Double, nY As Double, Nz As Double
  Call ClosestPointOnSegmentFromPointS3DXY(X1, Y1, z1, X2, Y2, z2, Px, Py, Pz, nX, nY, Nz)
  MinimumDistanceFromPointToSegment3DXY = Distance3DXY(Px, Py, Pz, nX, nY, Nz)
End Function
Public Function MinimumDistanceFromPointToSegment3D(Point As TPoint3D, Segment As TSegment3D) As Double
  MinimumDistanceFromPointToSegment3D = MinimumDistanceFromPointToSegment3DXY(Point.X, Point.Y, Point.z, Segment.P(1).X, Segment.P(1).Y, Segment.P(1).z, Segment.P(2).X, Segment.P(2).Y, Segment.P(2).z)
End Function
'(* Endof Minimum Distance From Point to Segment *)

'MinimumDistanceFromPointToLine
Public Function MinimumDistanceFromPointToLine2DXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
Dim nX As Double, nY As Double
  Call ClosestPointOnLineFromPointS2DXY(X1, Y1, X2, Y2, Px, Py, nX, nY)
  MinimumDistanceFromPointToLine2DXY = Distance2DXY(Px, Py, nX, nY)
End Function
Public Function MinimumDistanceFromPointToLine2DPL(Point As TPoint2D, Line As TLine2D) As Double
  MinimumDistanceFromPointToLine2DPL = MinimumDistanceFromPointToLine2DXY(Point.X, Point.Y, Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y)
End Function
Public Function MinimumDistanceFromPointToLine3DXY(Px As Double, Py As Double, Pz As Double, X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As Double
Dim nX As Double, nY As Double, Nz As Double
  Call ClosestPointOnLineFromPointS3DXY(X1, Y1, z1, X2, Y2, z2, Px, Py, Pz, nX, nY, Nz)
  MinimumDistanceFromPointToLine3DXY = Distance3DXY(Px, Py, Pz, nX, nY, Nz)
End Function
Public Function MinimumDistanceFromPointToLine3DPL(Point As TPoint3D, Line As TLine3D) As Double
  MinimumDistanceFromPointToLine3DPL = MinimumDistanceFromPointToLine3DXY(Point.X, Point.Y, Point.z, Line.P(0).X, Line.P(0).Y, Line.P(0).z, Line.P(1).X, Line.P(1).Y, Line.P(1).z)
End Function
'(* Endof Minimum Distance From Point to Line *)

Public Function MinimumDistanceFromPointToTriangleXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Double
Dim nX As Double, nY As Double
  Call ClosestPointOnTriangleFromPointS2DXY(X1, Y1, X2, Y2, X3, Y3, Px, Py, nX, nY)
  MinimumDistanceFromPointToTriangleXY = Distance2DXY(Px, Py, nX, nY)
End Function
Public Function MinimumDistanceFromPointToTriangle(Point As TPoint2D, Triangle As TTriangle2D) As Double
  MinimumDistanceFromPointToTriangle = MinimumDistanceFromPointToTriangleXY(Point.X, Point.Y, Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y)
End Function
'(* Endof Minimum Distance From Point to Triangle *)


'MinimumDistanceFromPointToPolygon
Public Function MinimumDistanceFromPointToRectangleXY(Px As Double, Py As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
Dim nX As Double, nY As Double
  Call ClosestPointOnRectangleFromPointXY(X1, Y1, X2, Y2, Px, Py, nX, nY)
  MinimumDistanceFromPointToRectangleXY = Distance2DXY(Px, Py, nX, nY)
End Function
Public Function MinimumDistanceFromPointToRectangle(Point As TPoint2D, aRectangle As TRectangle) As Double
  MinimumDistanceFromPointToRectangle = MinimumDistanceFromPointToRectangleXY(Point.X, Point.Y, aRectangle.P(0).X, aRectangle.P(0).Y, aRectangle.P(1).X, aRectangle.P(1).Y)
End Function
'(* Endof Minimum Distance From Point to Rectangle *)

Public Function MinimumDistanceFromPointToPolygon(Point As TPoint2D, Polygon As TPolygon2D) As Double
Dim i As Long, j As Long, TempDist As Double
  MinimumDistanceFromPointToPolygon = 0#
  If UBound(Polygon.Arr) < 3 Then Exit Function
  j = UBound(Polygon.Arr)
  i = 0
  MinimumDistanceFromPointToPolygon = MinimumDistanceFromPointToSegment2DXY(Point.X, Point.Y, Polygon.Arr(i).X, Polygon.Arr(i).Y, Polygon.Arr(j).X, Polygon.Arr(j).Y)
  j = 0
  For i = 1 To UBound(Polygon.Arr)
    TempDist = MinimumDistanceFromPointToSegment2DXY(Point.X, Point.Y, Polygon.Arr(i).X, Polygon.Arr(i).Y, Polygon.Arr(j).X, Polygon.Arr(j).Y)
    If TempDist < MinimumDistanceFromPointToPolygon Then
      MinimumDistanceFromPointToPolygon = TempDist
    End If
    j = i
  Next
End Function
'(* Endof Minimum Distance From Point to Polygon *)

'SegmentMidPoint
'SegmentMidPoint
Public Sub SegmentMidPointS2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, midx As Double, midy As Double)
  midx = (X1 + X2) * 0.5
  midy = (Y1 + Y2) * 0.5
End Sub
Public Sub SegmentMidPointS2DS(Segment As TSegment2D, midx As Double, midy As Double)
  Call SegmentMidPointS2DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, midx, midy)
End Sub
Public Function SegmentMidPoint2DP(P1 As TPoint2D, P2 As TPoint2D) As TPoint2D
  Call SegmentMidPointS2DXY(P1.X, P1.Y, P2.X, P2.Y, SegmentMidPoint2DP.X, SegmentMidPoint2DP.Y)
End Function
Public Function SegmentMidPoint2DS(Segment As TSegment2D) As TPoint2D
  SegmentMidPoint2DS = SegmentMidPoint2DP(Segment.P(0), Segment.P(1))
End Function
Public Sub SegmentMidPointS3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, midx As Double, midy As Double, midz As Double)
  midx = (X1 + X2) * 0.5
  midy = (Y1 + Y2) * 0.5
  midz = (z1 + z2) * 0.5
End Sub
Public Function SegmentMidPoint3DP(P1 As TPoint3D, P2 As TPoint3D) As TPoint3D
  Call SegmentMidPointS3DXY(P1.X, P1.Y, P1.z, P2.X, P2.Y, P2.z, SegmentMidPoint3DP.X, SegmentMidPoint3DP.Y, SegmentMidPoint3DP.z)
End Function
Public Function SegmentMidPoint3DS(Segment As TSegment3D) As TPoint3D
  SegmentMidPoint3DS = SegmentMidPoint3DP(Segment.P(0), Segment.P(1))
End Function
'(* Endof SegmentMidPoint *)

'Centroid
Public Sub CentroidS2D2XY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X As Double, Y As Double)
  X = (X1 + X2) * 0.5
  Y = (Y1 + Y2) * 0.5
End Sub
Public Function Centroid2DP(P1 As TPoint2D, P2 As TPoint2D) As TPoint2D
  Call CentroidS2D2XY(P1.X, P1.Y, P2.X, P2.Y, Centroid2DP.X, Centroid2DP.Y)
End Function
Public Function Centroid2DS(Segment As TSegment2D) As TPoint2D
  Centroid2DS = Centroid2DP(Segment.P(0), Segment.P(1))
End Function
Public Sub CentroidS2D3XY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X As Double, Y As Double)
Dim midx1 As Double, midy1 As Double, midx2 As Double, midy2 As Double
  Call SegmentMidPointS2DXY(X2, Y2, X3, Y3, midx1, midy1)
  Call SegmentMidPointS2DXY(X1, Y1, X3, Y3, midx2, midy2)
  Call Intersect2DXYi(X1, Y1, midx1, midy1, X2, Y2, midx2, midy2, X, Y)
End Sub
Public Sub CentroidS2DT(Triangle As TTriangle2D, X As Double, Y As Double)
  Call CentroidS2D3XY(Triangle.P(0).X, Triangle.P(0).Y, _
                      Triangle.P(1).X, Triangle.P(1).Y, _
                      Triangle.P(2).X, Triangle.P(2).Y, X, Y)
End Sub
Public Sub CentroidS2DR(Rectangle As TRectangle, X As Double, Y As Double)
  X = (Rectangle.P(0).X + Rectangle.P(1).X) * 0.5
  X = (Rectangle.P(0).Y + Rectangle.P(1).Y) * 0.5
End Sub
Public Function Centroid2D3P(P1 As TPoint2D, P2 As TPoint2D, P3 As TPoint2D) As TPoint2D
  Call CentroidS2D3XY(P1.X, P1.Y, P2.X, P2.Y, P3.X, P3.Y, Centroid2D3P.X, Centroid2D3P.Y)
End Function
Public Function Centroid2DT(Triangle As TTriangle2D) As TPoint2D
  Centroid2DT = Centroid2D3P(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Function
Public Function Centroid2DR(Rectangle As TRectangle) As TPoint2D
  Call CentroidS2DR(Rectangle, Centroid2DR.X, Centroid2DR.Y)
End Function
Public Sub CentroidS2DPg(Polygon As TPolygon2D, X As Double, Y As Double)
Dim i As Long, j As Long, asum As Double, term As Double
  X = 0#
  Y = 0#
  If UBound(Polygon.Arr) < 2 Then Exit Sub
  asum = 0#
  j = UBound(Polygon.Arr)
  For i = 0 To UBound(Polygon.Arr)
    term = ((Polygon.Arr(j).X * Polygon.Arr(i).Y) - (Polygon.Arr(j).Y * Polygon.Arr(i).X))
    asum = asum + term
    X = X + (Polygon.Arr(j).X + Polygon.Arr(i).X) * term
    Y = Y + (Polygon.Arr(j).Y + Polygon.Arr(i).Y) * term
    j = i
  Next
  If NotEqual(asum, 0#) Then
    X = X / (3# * asum)
    Y = Y / (3# * asum)
  End If
End Sub
Public Function Centroid2DPg(Polygon As TPolygon2D) As TPoint2D
  Call CentroidS2DPg(Polygon, Centroid2DPg.X, Centroid2DPg.Y)
End Function
Public Function Centroid3DPA(Polygon() As TPoint3D) As TPoint3D
Dim TotalArea As Double, TempArea  As Double, i As Long, lLen As Long
  Centroid3DPA.X = 0#
  Centroid3DPA.Y = 0#
  Centroid3DPA.z = 0#
  If UBound(Polygon) < 2 Then Exit Function
  TotalArea = 0#
  lLen = UBound(Polygon)
  For i = 0 To lLen - (2 - 1)
    TempArea = Area3DP(Polygon(i), Polygon(i + 1), Polygon(lLen))
    TotalArea = TotalArea + TempArea
    With Centroid3DPA
      .X = .X + TempArea * (Polygon(i).X + Polygon(i + 1).X + Polygon(lLen).X) / 3#
      .Y = .Y + TempArea * (Polygon(i).Y + Polygon(i + 1).Y + Polygon(lLen).Y) / 3#
      .z = .z + TempArea * (Polygon(i).z + Polygon(i + 1).z + Polygon(lLen).z) / 3#
    End With
  Next
  With Centroid3DPA
    .X = .X / TotalArea
    .Y = .Y / TotalArea
    .z = .z / TotalArea
  End With
End Function
'(* Endof Centroid *)

'PolygonSegmentIntersect
Public Function PolygonSegmentIntersect(Segment As TSegment2D, Polygon As TPolygon2D) As Boolean
Dim i As Long, j As Long
  PolygonSegmentIntersect = False
  If UBound(Polygon.Arr) < 3 Then Exit Function
  j = UBound(Polygon.Arr)
  For i = 0 To UBound(Polygon.Arr)
    If Intersect2DP(Segment.P(0), Segment.P(1), Polygon.Arr(i), Polygon.Arr(j)) Then
      PolygonSegmentIntersect = True
      Exit Function
    End If
    j = i
  Next
End Function
'(* Endof PolygonSegmentIntersect *)

'PolygonInPolygon
Public Function PolygonInPolygon(Poly1, Poly2 As TPolygon2D) As Boolean
  '(* to be implemented at a later date *)
  PolygonInPolygon = False
End Function
'(* Endof PolygonInPolygon *)

'PointInConvexPolygon
Public Function PointInConvexPolygonXY(Px As Double, Py As Double, Polygon As TPolygon2D) As Boolean
Dim i As Long, j As Long, InitialOrientation As Long
  PointInConvexPolygonXY = False
  If UBound(Polygon.Arr) < 3 Then Exit Function
  PointInConvexPolygonXY = True
  InitialOrientation = Orientation2DPXY(Polygon.Arr(0), Polygon.Arr(UBound(Polygon.Arr)), Px, Py) 'UBound(polygon.Arr)-1 ??
  j = 0
  If InitialOrientation <> 0 Then
    For i = 1 To UBound(Polygon.Arr)
      If InitialOrientation <> Orientation2DPXY(Polygon.Arr(i), Polygon.Arr(j), Px, Py) Then
        PointInConvexPolygonXY = False
        Exit Function
      End If
      j = i
    Next
  End If
End Function
Public Function PointInConvexPolygon(Point As TPoint2D, Polygon As TPolygon2D) As Boolean
  PointInConvexPolygon = PointInConvexPolygonXY(Point.X, Point.Y, Polygon)
End Function
'(* Endof PointInConvexPolygon *)

'PointInConcavePolygon
Public Function PointInConcavePolygonXY(Px As Double, Py As Double, Polygon As TPolygon2D) As Boolean
 '(* to be implemented at a later date *)
  PointInConcavePolygonXY = False
End Function
Public Function PointInConcavePolygon(Point As TPoint2D, Polygon As TPolygon2D) As Boolean
  PointInConcavePolygon = PointInConcavePolygonXY(Point.X, Point.Y, Polygon)
End Function
'(* Endof PointInConcavePolygon *)

'PointOnPolygon
Public Function PointOnPolygonXY(Px As Double, Py As Double, Polygon As TPolygon2D) As Boolean
Dim i As Long, j As Long
  PointOnPolygonXY = False
  If UBound(Polygon.Arr) < 3 Then Exit Function
  j = UBound(Polygon.Arr)
  For i = 0 To UBound(Polygon.Arr)
    If ((Polygon.Arr(i).Y <= Py) And (Py < Polygon.Arr(j).Y)) Or _
       ((Polygon.Arr(j).Y <= Py) And (Py < Polygon.Arr(i).Y)) Then
      If IsPointCollinear2DPXY(Polygon.Arr(i), Polygon.Arr(j), Px, Py) Then
        PointOnPolygonXY = True
        Exit Function
      End If
    End If
    j = i
  Next
End Function
Public Function PointOnPolygon(Point As TPoint2D, Polygon As TPolygon2D) As Boolean
  PointOnPolygon = PointOnPolygonXY(Point.X, Point.Y, Polygon)
End Function
'(* Endof PointOnPolygon *)

'PointInPolygon
Public Function PointInPolygonXY(Px As Double, Py As Double, Polygon As TPolygon2D) As Boolean
Dim i As Long, j As Long
  PointInPolygonXY = False
  If UBound(Polygon.Arr) < 3 Then Exit Function
  j = UBound(Polygon.Arr)
  For i = 0 To UBound(Polygon.Arr)                            '// an upward crossing
    If ((Polygon.Arr(i).Y <= Py) And (Py < Polygon.Arr(j).Y)) Or _
       ((Polygon.Arr(j).Y <= Py) And (Py < Polygon.Arr(i).Y)) Then  '// a downward crossing
      '(* compute the edge-ray intersect @ the x-coordinate *)
      If (Px - Polygon.Arr(i).X < ((Polygon.Arr(j).X - Polygon.Arr(i).X) * (Py - Polygon.Arr(i).Y) / (Polygon.Arr(j).Y - Polygon.Arr(i).Y))) Then
        PointInPolygonXY = Not PointInPolygonXY
      End If
    End If
    j = i
  Next
End Function
Public Function PointInPolygon(Point As TPoint2D, Polygon As TPolygon2D) As Boolean
  PointInPolygon = PointInPolygonXY(Point.X, Point.Y, Polygon)
End Function
Public Function PointInPolygon2DA(Point() As TPoint2D, Polygon As TPolygon2D) As TBooleanArray
Dim i As Long, j As Long, k As Long, Ratio As Double, RatioCalculated As Boolean
  Ratio = 0
  RatioCalculated = False
  If (UBound(Polygon.Arr) < 3) Or (UBound(Point) = 0) Then Exit Function
  j = UBound(Polygon.Arr)
  ReDim PointInPolygon2DA.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(PointInPolygon2DA.Arr):   PointInPolygon2DA.Arr(i) = False: Next
  For i = 0 To UBound(Polygon.Arr)
    For k = 0 To UBound(Point)                                                            '// an upward crossing
      If ((Polygon.Arr(i).Y <= Point(k).Y) And (Point(k).Y < Polygon.Arr(j).Y)) Or _
         ((Polygon.Arr(j).Y <= Point(k).Y) And (Point(k).Y < Polygon.Arr(i).Y)) Then  '// a wnward crossing
        If Not RatioCalculated Then
          Ratio = (Polygon.Arr(j).X - Polygon.Arr(i).X) / (Polygon.Arr(j).Y - Polygon.Arr(i).Y)
          RatioCalculated = True
        End If
        If (Point(k).X - Polygon.Arr(i).X < ((Point(k).Y - Polygon.Arr(i).Y) * Ratio)) Then
          PointInPolygon2DA.Arr(k) = Not PointInPolygon2DA.Arr(k)
        End If
      End If
    Next
    RatioCalculated = False
    j = i
  Next
End Function
'(* End PointInPolygon *)

'ConvexQuadix
Public Function ConvexQuadix(Quadix As TQuadix2D) As Boolean
Dim Orin As Double
  ConvexQuadix = False
      Orin = Orientation2DP(Quadix.P(0), Quadix.P(2), Quadix.P(1))
  If Orin <> Orientation2DP(Quadix.P(1), Quadix.P(3), Quadix.P(2)) Then Exit Function
  If Orin <> Orientation2DP(Quadix.P(2), Quadix.P(0), Quadix.P(3)) Then Exit Function
  If Orin <> Orientation2DP(Quadix.P(3), Quadix.P(1), Quadix.P(0)) Then Exit Function
  ConvexQuadix = True
End Function
'(* Endof ConvexQuadix *)

'ComplexPolygon
Public Function ComplexPolygon(Polygon As TPolygon2D) As Boolean
  '(*
  '  Complex polygon definition
  '  A polygon that can have:
  '  1.) Self intersecting edges
  '  2.) Holes
  '  3.) Unclosed area
  '*)
  ComplexPolygon = SelfIntersectingPolygon(Polygon)
End Function
'(* Endof ComplexPolygon *)

'SimplePolygon
Public Function SimplePolygon(Polygon As TPolygon2D) As Boolean
  '(*
  '  Simple polygon definition
  '  A polygon that can have:
  '  1.) inner and outter verticies
  '  2.) closed area
  '  3.) no self intersecting edges
  '*)
  SimplePolygon = Not SelfIntersectingPolygon(Polygon)
End Function
'(* Endof SimplePolygon *)

'ConvexPolygon
Public Function ConvexPolygon(Polygon As TPolygon2D) As Boolean
Dim i As Long, j As Long, k As Long, InitialOrientation As Long, CurrentOrientation As Long, FirstTime As Boolean
  ConvexPolygon = False
  If UBound(Polygon.Arr) < 3 Then Exit Function
  FirstTime = True
  InitialOrientation = Orientation2DPXY(Polygon.Arr(UBound(Polygon.Arr) - 1), Polygon.Arr(UBound(Polygon.Arr)), Polygon.Arr(0).X, Polygon.Arr(0).Y)
  j = 0
  k = UBound(Polygon.Arr)
  For i = 1 To UBound(Polygon.Arr)
    CurrentOrientation = Orientation2DPXY(Polygon.Arr(k), Polygon.Arr(j), Polygon.Arr(i).X, Polygon.Arr(i).Y)
    If (InitialOrientation = CollinearOrientation) And (InitialOrientation <> CurrentOrientation) And FirstTime Then
      InitialOrientation = CurrentOrientation
      FirstTime = False
    ElseIf (InitialOrientation <> CurrentOrientation) And (CurrentOrientation <> CollinearOrientation) Then
      Exit Function
    End If
    k = j
    j = i
  Next
  ConvexPolygon = True
End Function
Public Function ConvexPolygonA(Polygon() As TPoint2D) As Boolean
Dim i As Long, j As Long, k As Long, InitialOrientation As Long, CurrentOrientation As Long
  ConvexPolygonA = False
  If UBound(Polygon) < 3 Then Exit Function
  InitialOrientation = Orientation2DPXY(Polygon(UBound(Polygon) - 1), Polygon(UBound(Polygon)), Polygon(0).X, Polygon(0).Y)
  j = 0
  k = UBound(Polygon)
  For i = 1 To UBound(Polygon)
    CurrentOrientation = Orientation2DPXY(Polygon(k), Polygon(j), Polygon(i).X, Polygon(i).Y)
    If (InitialOrientation = CollinearOrientation) And (InitialOrientation <> CurrentOrientation) Then
      InitialOrientation = CurrentOrientation
    ElseIf (InitialOrientation <> CurrentOrientation) And (CurrentOrientation <> CollinearOrientation) Then
      Exit Function
    End If
    k = j
    j = i
  Next
  ConvexPolygonA = True
End Function
'(* Endof ConvexPolygon *)

'ConcavePolygon
Public Function ConcavePolygon(Polygon As TPolygon2D) As Boolean
  ConcavePolygon = Not ConvexPolygon(Polygon)
End Function
'(* Endof ConcavePolygon *)

'ConvexPolygonOrientation
Public Function ConvexPolygonOrientation(Polygon As TPolygon2D) As Long
  If UBound(Polygon.Arr) < 3 Then
    ConvexPolygonOrientation = 0
  Else
    ConvexPolygonOrientation = Orientation2DP(Polygon.Arr(0), Polygon.Arr(1), Polygon.Arr(2))
  End If
End Function
'(* Endof ConvexPolygonOrientation *)

'SimplePolygonOrientation
Public Function SimplePolygonOrientation(Polygon As TPolygon2D) As Long
Dim i As Long, Anchor  As Long, prevpos As Long, postpos As Long
  SimplePolygonOrientation = 0
  If UBound(Polygon.Arr) < 3 Then Exit Function
  Anchor = 0
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).X > Polygon.Arr(Anchor).X Then
      Anchor = i
    ElseIf (Polygon.Arr(i).X = Polygon.Arr(Anchor).X) And (Polygon.Arr(i).Y < Polygon.Arr(Anchor).Y) Then
      Anchor = i
    End If
  Next
  postpos = (Anchor + 1) Mod UBound(Polygon.Arr)
  prevpos = Anchor - 1
  If prevpos < 0 Then
    prevpos = UBound(Polygon.Arr) - prevpos
  End If
  SimplePolygonOrientation = Orientation2DP(Polygon.Arr(prevpos), Polygon.Arr(postpos), Polygon.Arr(Anchor))
End Function
'(* Endof SimplePolygonOrientation *)

'SelfIntersectingPolygon
Public Function SelfIntersectingPolygon(Polygon As TPolygon2D) As Boolean
Dim i As Long, j As Long, Poly1Trailer As Long, Poly2Trailer As Long
  SelfIntersectingPolygon = False ' Or (UBound(Polygon.arr) < 3)
  If (UBound(Polygon.Arr) < 3) Then Exit Function
  Poly1Trailer = UBound(Polygon.Arr)
  For i = 0 To UBound(Polygon.Arr)
    Poly2Trailer = i + 1
    For j = i + 2 To UBound(Polygon.Arr) - 1 ' - 2
      If (i <> j) And (Poly1Trailer <> Poly2Trailer) Then
        If Intersect2DP(Polygon.Arr(i), Polygon.Arr(Poly1Trailer), Polygon.Arr(j), Polygon.Arr(Poly2Trailer)) Then
          SelfIntersectingPolygon = True
          Exit Function
        End If
      End If
      Poly2Trailer = j
    Next
    Poly1Trailer = i
  Next
End Function
'(* Endof SelfIntersectingPolygon *)

'RectangularHull
Public Function RectangularHullA(Point() As TPoint2D) As TRectangle
Dim MaxX As Double, MaxY As Double, MinX As Double, MinY As Double, i As Long
  If (UBound(Point) + 1) < 2 Then Exit Function
  MinX = Point(0).X
  MaxX = Point(0).X
  MinY = Point(0).Y
  MaxY = Point(0).Y
  For i = 1 To UBound(Point)
    If Point(i).X < MinX Then
      MinX = Point(i).X
    ElseIf Point(i).X > MaxX Then
      MaxX = Point(i).X
    End If
    If Point(i).Y < MinY Then
      MinY = Point(i).Y
    ElseIf Point(i).Y > MaxY Then
      MaxY = Point(i).Y
    End If
  Next
  RectangularHullA = EquateRectangleXY(MinX, MinY, MaxX, MaxY)
End Function
Public Function RectangularHull(Polygon As TPolygon2D) As TRectangle
Dim MaxX As Double, MaxY As Double, MinX As Double, MinY As Double, i As Long
  If UBound(Polygon.Arr) < 2 Then Exit Function
  MinX = Polygon.Arr(0).X
  MaxX = Polygon.Arr(0).X
  MinY = Polygon.Arr(0).Y
  MaxY = Polygon.Arr(0).Y
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).X < MinX Then
      MinX = Polygon.Arr(i).X
    ElseIf Polygon.Arr(i).X > MaxX Then
      MaxX = Polygon.Arr(i).X
    End If
    If Polygon.Arr(i).Y < MinY Then
      MinY = Polygon.Arr(i).Y
    ElseIf Polygon.Arr(i).Y > MaxY Then
      MaxY = Polygon.Arr(i).Y
    End If
  Next
  RectangularHull = EquateRectangleXY(MinX, MinY, MaxX, MaxY)
End Function
'(* Endof RectangularHull *)

'CircularHull
Public Function CircularHull(Polygon As TPolygon2D) As TCircle
Dim i As Long, Cen As TPoint2D, LayLen  As Double, LayDist As Double
  If UBound(Polygon.Arr) < 3 Then Exit Function
  LayLen = -1
  Cen = Centroid2DPg(Polygon)
  For i = 0 To UBound(Polygon.Arr)
    LayDist = LayDistance2DPP(Cen, Polygon.Arr(i))
    If LayDist > LayLen Then
      LayLen = LayDist
    End If
  Next
  CircularHull.X = Cen.X
  CircularHull.Y = Cen.Y
  CircularHull.Radius = Sqr(LayLen)
End Function
'(* Endof CircularHull *)

'SphereHull
Public Function SphereHull(Polygon() As TPoint3D) As TSphere
Dim i As Long, Cen As TPoint3D, LayLen As Double, LayDist As Double
  If UBound(Polygon) < 2 Then Exit Function
  LayLen = -1
  Cen = Centroid3DPA(Polygon)
  For i = 0 To UBound(Polygon)
    LayDist = LayDistance3DPP(Cen, Polygon(i))
    If LayDist > LayLen Then
      LayLen = LayDist
    End If
  Next
  SphereHull.X = Cen.X
  SphereHull.Y = Cen.Y
  SphereHull.z = Cen.z
  SphereHull.Radius = Sqr(LayLen)
End Function
'(* Endof SphereHull *)

'CalculateBarycentricBase
Public Function CalculateBarycentricBase(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As Double
  CalculateBarycentricBase = Signed2D(X1, Y1, X2, Y2, X3, Y3)
End Function
'(* End Of Calculate Barycentric Unit *)

'CreateBarycentricUnit
'barycenter = Flächenschwerpunkt
Public Function CreateBarycentricUnit(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As TBarycentricUnit
  CreateBarycentricUnit.X1 = X1
  CreateBarycentricUnit.Y1 = Y1
  CreateBarycentricUnit.X2 = X2
  CreateBarycentricUnit.Y2 = Y2
  CreateBarycentricUnit.X3 = X3
  CreateBarycentricUnit.Y3 = Y3
  CreateBarycentricUnit.delta = CalculateBarycentricBase(X1, Y1, X2, Y2, X3, Y3)
End Function
Public Function CreateBarycentricUnitT(Triangle As TTriangle2D) As TBarycentricUnit
  CreateBarycentricUnitT = CreateBarycentricUnit(Triangle.P(0).X, Triangle.P(0).Y, _
                                                 Triangle.P(1).X, Triangle.P(1).Y, _
                                                 Triangle.P(2).X, Triangle.P(2).Y)
End Function
'(* Endof Create Barycentric Unit *)

'ConvertCartesianToBarycentric
Public Sub ConvertCartesianToBarycentricSXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Px As Double, Py As Double, u As Double, v As Double, W As Double)
Dim BarycentricBase As Double
  BarycentricBase = 1 / CalculateBarycentricBase(X1, Y1, X2, Y2, X3, Y3)
  u = CalculateBarycentricBase(Px, Py, X2, Y2, X3, Y3) * BarycentricBase
  v = CalculateBarycentricBase(X1, Y1, Px, Py, X3, Y3) * BarycentricBase
  W = CalculateBarycentricBase(X1, Y1, X2, Y2, Px, Py) * BarycentricBase
End Sub
Public Sub ConvertCartesianToBarycentricSUXY(BU As TBarycentricUnit, Px As Double, Py As Double, u As Double, v As Double, W As Double)
Dim BarycentricBase As Double
  BarycentricBase = 1 / BU.delta
  u = CalculateBarycentricBase(Px, Py, BU.X2, BU.Y2, BU.X3, BU.Y3) * BarycentricBase
  v = CalculateBarycentricBase(BU.X1, BU.Y1, Px, Py, BU.X3, BU.Y3) * BarycentricBase
  W = CalculateBarycentricBase(BU.X1, BU.Y1, BU.X2, BU.Y2, Px, Py) * BarycentricBase
End Sub
Public Sub ConvertCartesianToBarycentricSUPT(BU As TBarycentricUnit, Point As TPoint2D, BCrd As TBarycentricTriplet)
  Call ConvertCartesianToBarycentricSUXY(BU, Point.X, Point.Y, BCrd.u, BCrd.v, BCrd.W)
End Sub
Public Function ConvertCartesianToBarycentricUP(BU As TBarycentricUnit, Point As TPoint2D) As TBarycentricTriplet
  Call ConvertCartesianToBarycentricSUPT(BU, Point, ConvertCartesianToBarycentricUP)
End Function
'(* End Of Convert Cartesian to Barycentric *)

'ConvertBarycentricToCartesian
Public Sub ConvertBarycentricToCartesianSXY(u As Double, v As Double, W As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X As Double, Y As Double)
  X = u * X1 + v * X2 + W * X3
  Y = u * Y1 + v * Y2 + W * Y3
End Sub
Public Sub ConvertBarycentricToCartesianSXYU(u As Double, v As Double, W As Double, BU As TBarycentricUnit, X As Double, Y As Double)
  Call ConvertBarycentricToCartesianSXY(u, v, W, BU.X1, BU.Y1, BU.X2, BU.Y2, BU.X3, BU.Y3, X, Y)
End Sub
Public Sub ConvertBarycentricToCartesianSXYUP(u As Double, v As Double, W As Double, BU As TBarycentricUnit, Point As TPoint2D)
  Call ConvertBarycentricToCartesianSXY(u, v, W, BU.X1, BU.Y1, BU.X2, BU.Y2, BU.X3, BU.Y3, Point.X, Point.Y)
End Sub
Public Function ConvertBarycentricToCartesianXYU(u As Double, v As Double, W As Double, BU As TBarycentricUnit) As TPoint2D
  Call ConvertBarycentricToCartesianSXYUP(u, v, W, BU, ConvertBarycentricToCartesianXYU)
End Function
'(* End Of Convert Barycentric to Cartesian *)

'Clip
Public Function OutCode(Rect As TRectangle, X As Double, Y As Double) As Long
Const CLIP_BOTTOM = 1 'OK!!
Const CLIP_TOP = 2
Const CLIP_LEFT = 4
Const CLIP_RIGHT = 8
  OutCode = 0
  If Y < Rect.P(0).Y Then
    OutCode = OutCode Or CLIP_TOP
  ElseIf Y > Rect.P(1).Y Then
    OutCode = OutCode Or CLIP_BOTTOM
  End If
  If X < Rect.P(0).X Then
    OutCode = OutCode Or CLIP_LEFT
  ElseIf X > Rect.P(1).X Then
    OutCode = OutCode Or CLIP_RIGHT
  End If
End Function
Public Function Clip2DSR(Segment As TSegment2D, Rect As TRectangle, CSegment As TSegment2D) As Boolean
 '(* Cohen-Sutherland Clipping Algorithm *)
Const CLIP_BOTTOM = 1
Const CLIP_TOP = 2
Const CLIP_LEFT = 4
Const CLIP_RIGHT = 8
Dim outcode0 As Long, outcode1 As Long, outcodeout As Long, X As Double, Y As Double, Dx As Double, Dy As Double
  
  Clip2DSR = False
  CSegment = Segment
  X = 0#
  Y = 0#
  
  outcode0 = OutCode(Rect, CSegment.P(0).X, CSegment.P(0).Y)
  outcode1 = OutCode(Rect, CSegment.P(1).X, CSegment.P(1).Y)
  
  Do While (outcode0 <> 0) Or (outcode1 <> 0)
    If (outcode0 And outcode1) <> 0 Then
      Exit Function
    Else
      If outcode0 <> 0 Then
        outcodeout = outcode0
      Else
        outcodeout = outcode1
      End If
      Dx = (CSegment.P(1).X - CSegment.P(0).X)
      Dy = (CSegment.P(1).Y - CSegment.P(0).Y)
      
      If ((outcodeout And CLIP_TOP) = CLIP_TOP) Then
        X = CSegment.P(0).X + Dx * (Rect.P(0).Y - CSegment.P(0).Y) / Dy
        Y = Rect.P(0).Y
      ElseIf ((outcodeout And CLIP_BOTTOM) = CLIP_BOTTOM) Then
        X = CSegment.P(0).X + Dx * (Rect.P(1).Y - CSegment.P(0).Y) / Dy
        Y = Rect.P(1).Y
      ElseIf ((outcodeout And CLIP_RIGHT) = CLIP_RIGHT) Then
        Y = CSegment.P(0).Y + Dy * (Rect.P(1).X - CSegment.P(0).X) / Dx
        X = Rect.P(1).X
      ElseIf ((outcodeout And CLIP_LEFT) = CLIP_LEFT) Then
        Y = CSegment.P(0).Y + Dy * (Rect.P(0).X - CSegment.P(0).X) / Dx
        X = Rect.P(0).X
      End If
      If (outcodeout = outcode0) Then
        CSegment.P(0).X = X
        CSegment.P(0).Y = Y
        outcode0 = OutCode(Rect, CSegment.P(0).X, CSegment.P(0).Y)
      Else
        CSegment.P(1).X = X
        CSegment.P(1).Y = Y
        outcode1 = OutCode(Rect, CSegment.P(1).X, CSegment.P(1).Y)
      End If
    End If
  Loop
  Clip2DSR = True
End Function
Public Function Clip2DST(Segment As TSegment2D, Triangle As TTriangle2D, CSegment As TSegment2D) As Boolean
Dim Pos As Long
  Pos = 1
  If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Triangle.P(0).X, Triangle.P(0).Y, Triangle.P(1).X, Triangle.P(1).Y, CSegment.P(0).X, CSegment.P(0).Y) Then
    Call Inc(Pos)
  End If
  If Pos = 1 Then
    If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, CSegment.P(0).X, CSegment.P(0).Y) Then
      Call Inc(Pos)
    End If
  Else 'Pos = 2
    If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Triangle.P(1).X, Triangle.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, CSegment.P(1).X, CSegment.P(1).Y) Then
      Call Inc(Pos)
    End If
  End If
  If (Pos < 3) Then
    If Pos = 1 Then
      If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, Triangle.P(0).X, Triangle.P(0).Y, CSegment.P(0).X, CSegment.P(0).Y) Then
        Call Inc(Pos)
      End If
    Else 'Pos = 2
      If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Triangle.P(2).X, Triangle.P(2).Y, Triangle.P(0).X, Triangle.P(0).Y, CSegment.P(1).X, CSegment.P(1).Y) Then
        Call Inc(Pos)
      End If
    End If
  End If
  If Pos = 2 Then
    If PointInTriangle(Segment.P(0), Triangle) Then
      CSegment.P(1) = Segment.P(0)
    Else
      CSegment.P(1) = Segment.P(1)
    End If
  End If
  Clip2DST = (Pos > 1)
End Function
Public Function Clip2DSQ(Segment As TSegment2D, Quadix As TQuadix2D, CSegment As TSegment2D) As Boolean
Dim Pos As Long
  Pos = 1
  If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(1).X, Quadix.P(1).Y, CSegment.P(0).X, CSegment.P(0).Y) Then
    Call Inc(Pos)
  End If
  If Pos = 1 Then
    If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(2).X, Quadix.P(2).Y, CSegment.P(0).X, CSegment.P(0).Y) Then
      CSegment.P(0) = IntersectionPoint2DP(Segment.P(0), Segment.P(1), Quadix.P(1), Quadix.P(2))
      Call Inc(Pos)
    End If
  Else 'Pos = 2
    If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(2).X, Quadix.P(2).Y, CSegment.P(1).X, CSegment.P(1).Y) Then
      CSegment.P(1) = IntersectionPoint2DP(Segment.P(0), Segment.P(1), Quadix.P(1), Quadix.P(2))
      Call Inc(Pos)
    End If
  End If
  If (Pos < 3) Then
    If Pos = 1 Then
      If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Quadix.P(2).X, Quadix.P(2).Y, Quadix.P(3).X, Quadix.P(3).Y, CSegment.P(0).X, CSegment.P(0).Y) Then
        Call Inc(Pos)
      End If
    Else 'Pos = 2
      If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Quadix.P(2).X, Quadix.P(2).Y, Quadix.P(3).X, Quadix.P(3).Y, CSegment.P(1).X, CSegment.P(1).Y) Then
        Call Inc(Pos)
      End If
    End If
  End If
  If (Pos < 3) Then
    If Pos = 1 Then
      If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Quadix.P(3).X, Quadix.P(3).Y, Quadix.P(0).X, Quadix.P(0).Y, CSegment.P(0).X, CSegment.P(0).Y) Then
        Call Inc(Pos)
      End If
    Else 'Pos = 2
      If Intersect2DXYi(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y, Quadix.P(3).X, Quadix.P(3).Y, Quadix.P(0).X, Quadix.P(0).Y, CSegment.P(1).X, CSegment.P(1).Y) Then
        Call Inc(Pos)
      End If
    End If
  End If
  If Pos = 2 Then
    If PointInQuadix(Segment.P(0), Quadix) Then
      CSegment.P(1) = Segment.P(0)
    Else
      CSegment.P(1) = Segment.P(1)
    End If
  End If
  Clip2DSQ = (Pos > 1)
End Function
Public Function Clip2DSCi(Segment As TSegment2D, aCircle As TCircle, CSegment As TSegment2D) As Boolean
Dim Cnt As Long, I1  As TPoint2D, I2  As TPoint2D
  Clip2DSCi = False
  Call IntersectionPoint2DSCi(Segment, aCircle, Cnt, I1, I2)
  If Cnt = 2 Then
    CSegment.P(0) = I1
    CSegment.P(1) = I2
    Clip2DSCi = True
  End If
End Function
Public Function Clip2DSO(Segment As TSegment2D, Obj As TGeometricObject, CSegment As TSegment2D) As Boolean
  Clip2DSO = False
  Select Case Obj.ObjectType
  Case geoRectangle:  Clip2DSO = Clip2DSR(Segment, Obj.Rectangle, CSegment)
  Case geoTriangle2D: Clip2DSO = Clip2DST(Segment, Obj.Triangle2D, CSegment)
  Case geoQuadix2D:   Clip2DSO = Clip2DSQ(Segment, Obj.Quadix2D, CSegment)
  Case geoCircle:     Clip2DSO = Clip2DSCi(Segment, Obj.aCircle, CSegment)
  End Select
End Function
'(* End of Clip *)

'Area'OK!!
Public Function Area2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As Double
  Area2DP = 0.5 * ((Point1.X * (Point2.Y - Point3.Y)) + (Point2.X * (Point3.Y - Point1.Y)) + (Point3.X * (Point1.Y - Point2.Y)))
End Function
Public Function Area3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As Double
Dim Dx1 As Double, Dx2 As Double, Dy1 As Double, Dy2 As Double, Dz1 As Double, Dz2 As Double
Dim cX As Double, cY As Double, cz As Double 'OK!!
  Dx1 = Point2.X - Point1.X
  Dy1 = Point2.Y - Point1.Y
  Dz1 = Point2.z - Point1.z
  
  Dx2 = Point3.X - Point1.X
  Dy2 = Point3.Y - Point1.Y
  Dz2 = Point3.z - Point1.z
  
  cX = Dy1 * Dz2 - Dy2 * Dz1
  cY = Dx2 * Dz1 - Dx1 * Dz2
  cz = Dx1 * Dy2 - Dx2 * Dy1
  
  Area3DP = (Sqr(cX * cX + cY * cY + cz * cz) * 0.5)
End Function
Public Function Area2DT(Triangle As TTriangle2D) As Double 'OK!!
  Area2DT = 0.5 * ((Triangle.P(0).X * (Triangle.P(1).Y - Triangle.P(2).Y)) + _
                   (Triangle.P(1).X * (Triangle.P(2).Y - Triangle.P(0).Y)) + _
                   (Triangle.P(2).X * (Triangle.P(0).Y - Triangle.P(1).Y)))
End Function
Public Function Area3DT(Triangle As TTriangle3D) As Double 'OK!!
Dim Dx1 As Double, Dx2 As Double, Dy1 As Double, Dy2 As Double, Dz1 As Double, Dz2 As Double
Dim cX As Double, cY As Double, cz As Double
  Dx1 = Triangle.P(1).X - Triangle.P(0).X
  Dy1 = Triangle.P(1).Y - Triangle.P(0).Y
  Dz1 = Triangle.P(1).z - Triangle.P(0).z
  
  Dx2 = Triangle.P(2).X - Triangle.P(0).X
  Dy2 = Triangle.P(2).Y - Triangle.P(0).Y
  Dz2 = Triangle.P(2).z - Triangle.P(0).z
  
  cX = Dy1 * Dz2 - Dy2 * Dz1
  cY = Dx2 * Dz1 - Dx1 * Dz2
  cz = Dx1 * Dy2 - Dx2 * Dy1
  
  Area3DT = (Sqr(cX * cX + cY * cY + cz * cz) * 0.5)
End Function
Public Function Area2DQ(Quadix As TQuadix2D) As Double 'OK!!
  Area2DQ = 0.5 * ((Quadix.P(0).X * (Quadix.P(1).Y - Quadix.P(3).Y)) + _
                   (Quadix.P(1).X * (Quadix.P(2).Y - Quadix.P(0).Y)) + _
                   (Quadix.P(2).X * (Quadix.P(3).Y - Quadix.P(1).Y)) + _
                   (Quadix.P(3).X * (Quadix.P(0).Y - Quadix.P(2).Y)))
End Function
Public Function Area3DQ(Quadix As TQuadix3D) As Double 'OK!!
  Area3DQ = (Area3DT(EquateTriangle3DP(Quadix.P(0), Quadix.P(1), Quadix.P(2))) + _
             Area3DT(EquateTriangle3DP(Quadix.P(2), Quadix.P(3), Quadix.P(0))))
End Function
Public Function Area2DR(Rectangle As TRectangle) As Double 'OK!!
  Area2DR = Abs(Rectangle.P(1).X - Rectangle.P(0).X) * Abs(Rectangle.P(1).Y - Rectangle.P(0).Y)
End Function
Public Function Area2DCi(aCircle As TCircle) As Double 'OK!!
  Area2DCi = PI2 * aCircle.Radius * aCircle.Radius
End Function
Public Function Area2DPg(Polygon As TPolygon2D) As Double
Dim i As Long, j As Long 'OK!!
 '(*
 'Old implementation uses mod to wrap around - not very efficient...
 'Result = 0#
 'If Ubound(Polygon.Arr) < 3 Then Exit Function
 'For i = 0 To Ubound(Polygon.Arr) - 1
 '
 '  Result = Result + (
 '                      (Polygon.arr(i).x * Polygon[(i + 1) mod Ubound(Polygon.Arr)].y)-
 '                      (Polygon.arr(i).y * Polygon[(i + 1) mod Ubound(Polygon.Arr)].x)
 '                     )
 'Result = Result * 0.5
 '*)
  Area2DPg = 0#
  If UBound(Polygon.Arr) < 3 Then Exit Function
  j = UBound(Polygon.Arr)
  For i = 0 To UBound(Polygon.Arr)
    Area2DPg = Area2DPg + ((Polygon.Arr(j).X * Polygon.Arr(i).Y) - (Polygon.Arr(j).Y * Polygon.Arr(i).X))
    j = i
  Next
  Area2DPg = Area2DPg * 0.5
End Function
Public Function AreaO(Obj As TGeometricObject) As Double
  Select Case Obj.ObjectType 'OK!!
  Case geoTriangle2D: AreaO = Area2DT(Obj.Triangle2D)
  Case geoTriangle3D: AreaO = Area3DT(Obj.Triangle3D)
  Case geoQuadix2D:   AreaO = Area2DQ(Obj.Quadix2D)
  Case geoQuadix3D:   AreaO = Area3DQ(Obj.Quadix3D)
  Case geoRectangle:  AreaO = Area2DR(Obj.Rectangle)
  Case geoCircle:     AreaO = Area2DCi(Obj.aCircle)
  Case Else:          AreaO = 0#
  End Select
End Function
'(* Endof Area *)

'Perimeter
Public Function Perimeter2DT(Triangle As TTriangle2D) As Double 'OK!!
  Perimeter2DT = Distance2DP(Triangle.P(0), Triangle.P(1)) + _
                 Distance2DP(Triangle.P(1), Triangle.P(2)) + _
                 Distance2DP(Triangle.P(2), Triangle.P(0))
End Function
Public Function Perimeter3DT(Triangle As TTriangle3D) As Double 'OK!!
  Perimeter3DT = Distance3DP(Triangle.P(0), Triangle.P(1)) + _
                 Distance3DP(Triangle.P(1), Triangle.P(2)) + _
                 Distance3DP(Triangle.P(2), Triangle.P(0))
End Function
Public Function Perimeter2DQ(Quadix As TQuadix2D) As Double 'OK!!
  Perimeter2DQ = Distance2DP(Quadix.P(0), Quadix.P(1)) + _
                 Distance2DP(Quadix.P(1), Quadix.P(2)) + _
                 Distance2DP(Quadix.P(2), Quadix.P(3)) + _
                 Distance2DP(Quadix.P(3), Quadix.P(0))
End Function
Public Function Perimeter3DQ(Quadix As TQuadix3D) As Double 'OK!!
  Perimeter3DQ = Distance3DP(Quadix.P(0), Quadix.P(1)) + _
                 Distance3DP(Quadix.P(1), Quadix.P(2)) + _
                 Distance3DP(Quadix.P(2), Quadix.P(3)) + _
                 Distance3DP(Quadix.P(3), Quadix.P(0))
End Function
Public Function Perimeter2DR(Rectangle As TRectangle) As Double
  Perimeter2DR = 2 * (Abs(Rectangle.P(1).X - Rectangle.P(0).X) + Abs(Rectangle.P(1).Y - Rectangle.P(0).Y))
End Function
Public Function Perimeter2DCi(aCircle As TCircle) As Double
  Perimeter2DCi = PI2 * aCircle.Radius
End Function
Public Function Perimeter2DPg(Polygon As TPolygon2D) As Double
Dim i As Long, j As Long
  Perimeter2DPg = 0#
  If UBound(Polygon.Arr) < 3 Then Exit Function
  j = UBound(Polygon.Arr)
  For i = 0 To UBound(Polygon.Arr)
    Perimeter2DPg = Perimeter2DPg + Distance2DP(Polygon.Arr(i), Polygon.Arr(j))
    j = i
  Next
End Function
Public Function PerimeterO(Obj As TGeometricObject) As Double
  Select Case Obj.ObjectType
  Case geoTriangle2D: PerimeterO = Perimeter2DT(Obj.Triangle2D)
  Case geoTriangle3D: PerimeterO = Perimeter3DT(Obj.Triangle3D)
  Case geoQuadix2D:   PerimeterO = Perimeter2DQ(Obj.Quadix2D)
  Case geoQuadix3D:   PerimeterO = Perimeter3DQ(Obj.Quadix3D)
  Case geoRectangle:  PerimeterO = Perimeter2DR(Obj.Rectangle)
  Case geoCircle:     PerimeterO = Perimeter2DCi(Obj.aCircle)
  Case Else:          PerimeterO = 0#
  End Select
End Function
'(* Endof Perimeter *)

'SemiPerimeter
Public Function SemiPerimeter2D(Triangle As TTriangle2D) As Double
  SemiPerimeter2D = Perimeter2DT(Triangle) * 0.5
End Function
Public Function SemiPerimeter3D(Triangle As TTriangle3D) As Double
  SemiPerimeter3D = Perimeter3DT(Triangle) * 0.5
End Function
'(* Endof Perimeter *)

'Rotate
Public Sub RotateS2DXY(ByVal RotAng As Double, X As Double, Y As Double, nX As Double, nY As Double)
Dim SinVal As Double, CosVal As Double
  RotAng = RotAng * PIDiv180
  SinVal = Sin(RotAng)
  CosVal = Cos(RotAng)
  nX = (X * CosVal) - (Y * SinVal)
  nY = (Y * CosVal) + (X * SinVal)
End Sub
Public Sub RotateS2DXYo(RotAng As Double, X As Double, Y As Double, ox As Double, oy As Double, nX As Double, nY As Double)
  Call RotateS2DXY(RotAng, X - ox, Y - oy, nX, nY)
  nX = nX + ox
  nY = nY + oy
End Sub
Public Function Rotate2DP(RotAng As Double, Point As TPoint2D) As TPoint2D
  Call RotateS2DXY(RotAng, Point.X, Point.Y, Rotate2DP.X, Rotate2DP.Y)
End Function
Public Function Rotate2DPo(RotAng As Double, Point As TPoint2D, OPoint As TPoint2D) As TPoint2D
  Call RotateS2DXYo(RotAng, Point.X, Point.Y, OPoint.X, OPoint.Y, Rotate2DPo.X, Rotate2DPo.Y)
End Function
Public Function Rotate2DS(RotAng As Double, Segment As TSegment2D) As TSegment2D
  Rotate2DS.P(0) = Rotate2DP(RotAng, Segment.P(0)) 'OK!!
  Rotate2DS.P(1) = Rotate2DP(RotAng, Segment.P(1))
End Function
Public Function Rotate2DSo(RotAng As Double, Segment As TSegment2D, OPoint As TPoint2D) As TSegment2D
  Rotate2DSo.P(0) = Rotate2DPo(RotAng, Segment.P(0), OPoint) 'OK!!
  Rotate2DSo.P(1) = Rotate2DPo(RotAng, Segment.P(1), OPoint)
End Function
Public Function Rotate2DT(RotAng As Double, Triangle As TTriangle2D) As TTriangle2D
  Rotate2DT.P(0) = Rotate2DP(RotAng, Triangle.P(0)) 'OK!!
  Rotate2DT.P(1) = Rotate2DP(RotAng, Triangle.P(1))
  Rotate2DT.P(2) = Rotate2DP(RotAng, Triangle.P(2))
End Function
Public Function Rotate2DTo(RotAng As Double, Triangle As TTriangle2D, OPoint As TPoint2D) As TTriangle2D
  Rotate2DTo.P(0) = Rotate2DPo(RotAng, Triangle.P(0), OPoint) 'OK!!
  Rotate2DTo.P(1) = Rotate2DPo(RotAng, Triangle.P(1), OPoint)
  Rotate2DTo.P(2) = Rotate2DPo(RotAng, Triangle.P(2), OPoint)
End Function
Public Function Rotate2DQ(RotAng As Double, Quadix As TQuadix2D) As TQuadix2D
  Rotate2DQ.P(0) = Rotate2DP(RotAng, Quadix.P(0)) 'OK!!
  Rotate2DQ.P(1) = Rotate2DP(RotAng, Quadix.P(1))
  Rotate2DQ.P(2) = Rotate2DP(RotAng, Quadix.P(2))
  Rotate2DQ.P(3) = Rotate2DP(RotAng, Quadix.P(3))
End Function
Public Function Rotate2DQo(RotAng As Double, Quadix As TQuadix2D, OPoint As TPoint2D) As TQuadix2D
  Rotate2DQo.P(0) = Rotate2DPo(RotAng, Quadix.P(0), OPoint) 'OK!!
  Rotate2DQo.P(1) = Rotate2DPo(RotAng, Quadix.P(1), OPoint)
  Rotate2DQo.P(2) = Rotate2DPo(RotAng, Quadix.P(2), OPoint)
  Rotate2DQo.P(3) = Rotate2DPo(RotAng, Quadix.P(3), OPoint)
End Function
Public Function Rotate2DPg(RotAng As Double, Polygon As TPolygon2D) As TPolygon2D
Dim i As Long
  ReDim Rotate2DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    Rotate2DPg.Arr(i) = Rotate2DP(RotAng, Polygon.Arr(i))
  Next
End Function
Public Function Rotate2DPgo(RotAng As Double, Polygon As TPolygon2D, OPoint As TPoint2D) As TPolygon2D
Dim i As Long
  ReDim Rotate2DPgo.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    Rotate2DPgo.Arr(i) = Rotate2DPo(RotAng, Polygon.Arr(i), OPoint)
  Next
End Function
Public Function Rotate2DO(RotAng As Double, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint2D:       Rotate2DO.Point2D = Rotate2DP(RotAng, Obj.Point2D)
  Case geoSegment2D:   Rotate2DO.Segment2D = Rotate2DS(RotAng, Obj.Segment2D)
  Case geoTriangle2D: Rotate2DO.Triangle2D = Rotate2DT(RotAng, Obj.Triangle2D)
  Case geoQuadix2D:     Rotate2DO.Quadix2D = Rotate2DQ(RotAng, Obj.Quadix2D)
  Case Else:             Rotate2DO = Obj
  End Select
End Function
Public Function Rotate2DOo(RotAng As Double, Obj As TGeometricObject, OPoint As TPoint2D) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint2D:       Rotate2DOo.Point2D = Rotate2DPo(RotAng, Obj.Point2D, OPoint)
  Case geoSegment2D:   Rotate2DOo.Segment2D = Rotate2DSo(RotAng, Obj.Segment2D, OPoint)
  Case geoTriangle2D: Rotate2DOo.Triangle2D = Rotate2DTo(RotAng, Obj.Triangle2D, OPoint)
  Case geoQuadix2D:     Rotate2DOo.Quadix2D = Rotate2DQo(RotAng, Obj.Quadix2D, OPoint)
  Case Else:             Rotate2DOo = Obj
  End Select
End Function
Public Sub RotateS3D(Rx As Double, Ry As Double, Rz As Double, X As Double, Y As Double, z As Double, nX As Double, nY As Double, Nz As Double)
Dim TempX As Double, TempY As Double, TempZ As Double
Dim SinX As Double, SinY As Double, SinZ As Double
Dim CosX As Double, CosY As Double, CosZ As Double
Dim XRadAng As Double, YRadAng As Double, ZRadAng As Double
  XRadAng = Rx * PIDiv180
  YRadAng = Ry * PIDiv180
  ZRadAng = Rz * PIDiv180
  SinX = Sin(XRadAng)
  SinY = Sin(YRadAng)
  SinZ = Sin(ZRadAng)
  CosX = Cos(XRadAng)
  CosY = Cos(YRadAng)
  CosZ = Cos(ZRadAng)
  TempY = Y * CosY - z * SinY
  TempZ = Y * SinY + z * CosY
  TempX = X * CosX - TempZ * SinX
  Nz = X * SinX + TempZ * CosX
  nX = TempX * CosZ - TempY * SinZ
  nY = TempX * SinZ + TempY * CosZ
End Sub
Public Sub RotateS3DoN(Rx As Double, Ry As Double, Rz As Double, X As Double, Y As Double, z As Double, ox As Double, oy As Double, oz As Double, nX As Double, nY As Double, Nz As Double)
  Call RotateS3D(Rx, Ry, Rz, X - ox, Y - oy, z - oz, nX, nY, Nz)
  nX = nX + ox
  nY = nY + oy
  Nz = Nz + oz
End Sub
Public Function Rotate3DP(Rx As Double, Ry As Double, Rz As Double, Point As TPoint3D) As TPoint3D
  Call RotateS3D(Rx, Ry, Rz, Point.X, Point.Y, Point.z, Rotate3DP.X, Rotate3DP.Y, Rotate3DP.z)
End Function
Public Function Rotate3DPo(Rx As Double, Ry As Double, Rz As Double, Point As TPoint3D, OPoint As TPoint3D) As TPoint3D
  Call RotateS3DoN(Rx, Ry, Rz, Point.X, Point.Y, Point.z, OPoint.X, OPoint.Y, OPoint.z, Rotate3DPo.X, Rotate3DPo.Y, Rotate3DPo.z)
End Function
Public Function Rotate3DS(Rx As Double, Ry As Double, Rz As Double, Segment As TSegment3D) As TSegment3D
  With Rotate3DS 'OK!!
    .P(0) = Rotate3DP(Rx, Ry, Rz, Segment.P(0))
    .P(1) = Rotate3DP(Rx, Ry, Rz, Segment.P(1))
  End With
End Function
Public Function Rotate3DSo(Rx As Double, Ry As Double, Rz As Double, Segment As TSegment3D, OPoint As TPoint3D) As TSegment3D
  With Rotate3DSo 'OK!!
    .P(0) = Rotate3DPo(Rx, Ry, Rz, Segment.P(0), OPoint)
    .P(1) = Rotate3DPo(Rx, Ry, Rz, Segment.P(1), OPoint)
  End With
End Function
Public Function Rotate3DT(Rx As Double, Ry As Double, Rz As Double, Triangle As TTriangle3D) As TTriangle3D
  With Rotate3DT 'OK!!
    .P(0) = Rotate3DP(Rx, Ry, Rz, Triangle.P(0))
    .P(1) = Rotate3DP(Rx, Ry, Rz, Triangle.P(1))
    .P(2) = Rotate3DP(Rx, Ry, Rz, Triangle.P(2))
  End With
End Function
Public Function Rotate3DTo(Rx As Double, Ry As Double, Rz As Double, Triangle As TTriangle3D, OPoint As TPoint3D) As TTriangle3D
  With Rotate3DTo 'OK!!
    .P(0) = Rotate3DPo(Rx, Ry, Rz, Triangle.P(0), OPoint)
    .P(1) = Rotate3DPo(Rx, Ry, Rz, Triangle.P(1), OPoint)
    .P(2) = Rotate3DPo(Rx, Ry, Rz, Triangle.P(2), OPoint)
  End With
End Function
Public Function Rotate3DQ(Rx As Double, Ry As Double, Rz As Double, Quadix As TQuadix3D) As TQuadix3D
  With Rotate3DQ 'OK!!
    .P(0) = Rotate3DP(Rx, Ry, Rz, Quadix.P(0))
    .P(1) = Rotate3DP(Rx, Ry, Rz, Quadix.P(1))
    .P(2) = Rotate3DP(Rx, Ry, Rz, Quadix.P(2))
    .P(3) = Rotate3DP(Rx, Ry, Rz, Quadix.P(3))
  End With
End Function
Public Function Rotate3DQo(Rx As Double, Ry As Double, Rz As Double, Quadix As TQuadix3D, OPoint As TPoint3D) As TQuadix3D
  With Rotate3DQo 'OK!!
    .P(0) = Rotate3DPo(Rx, Ry, Rz, Quadix.P(0), OPoint)
    .P(1) = Rotate3DPo(Rx, Ry, Rz, Quadix.P(1), OPoint)
    .P(2) = Rotate3DPo(Rx, Ry, Rz, Quadix.P(2), OPoint)
    .P(3) = Rotate3DPo(Rx, Ry, Rz, Quadix.P(3), OPoint)
  End With
End Function
Public Function Rotate3DPg(Rx As Double, Ry As Double, Rz As Double, Polygon As TPolygon3D) As TPolygon3D
Dim i As Long 'OK!!
  ReDim Rotate3DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    Rotate3DPg.Arr(i) = Rotate3DP(Rx, Ry, Rz, Polygon.Arr(i))
  Next
End Function
Public Function Rotate3DPgo(Rx As Double, Ry As Double, Rz As Double, Polygon As TPolygon3D, OPoint As TPoint3D) As TPolygon3D
Dim i As Long 'OK!!
  ReDim Rotate3DPgo.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    Rotate3DPgo.Arr(i) = Rotate3DPo(Rx, Ry, Rz, Polygon.Arr(i), OPoint)
  Next
End Function
Public Function Rotate3DO(Rx As Double, Ry As Double, Rz As Double, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType 'OK!!
  Case geoPoint3D:       Rotate3DO.Point3D = Rotate3DP(Rx, Ry, Rz, Obj.Point3D)
  Case geoSegment3D:   Rotate3DO.Segment3D = Rotate3DS(Rx, Ry, Rz, Obj.Segment3D)
  Case geoTriangle3D: Rotate3DO.Triangle3D = Rotate3DT(Rx, Ry, Rz, Obj.Triangle3D)
  Case geoQuadix3D:     Rotate3DO.Quadix3D = Rotate3DQ(Rx, Ry, Rz, Obj.Quadix3D)
  Case Else:              Rotate3DO = Obj
  End Select
End Function
Public Function Rotate3DOo(Rx As Double, Ry As Double, Rz As Double, Obj As TGeometricObject, OPoint As TPoint3D) As TGeometricObject
  Select Case Obj.ObjectType 'OK!!
  Case geoPoint3D:       Rotate3DOo.Point3D = Rotate3DPo(Rx, Ry, Rz, Obj.Point3D, OPoint)
  Case geoSegment3D:   Rotate3DOo.Segment3D = Rotate3DSo(Rx, Ry, Rz, Obj.Segment3D, OPoint)
  Case geoTriangle3D: Rotate3DOo.Triangle3D = Rotate3DTo(Rx, Ry, Rz, Obj.Triangle3D, OPoint)
  Case geoQuadix3D:     Rotate3DOo.Quadix3D = Rotate3DQo(Rx, Ry, Rz, Obj.Quadix3D, OPoint)
  Case Else:          Rotate3DOo = Obj
  End Select
End Function
'(* Endof Rotate Ab,  Origin Point *)

'FastRotate
Public Sub FastRotateS2DXY(RotAng As Long, X As Double, Y As Double, nX As Double, nY As Double)
Dim SinVal As Double, CosVal As Double
  RotAng = RotAng Mod 360
  If RotAng < 0 Then RotAng = 360 + RotAng
  SinVal = SinTable(RotAng)
  CosVal = CosTable(RotAng)
  nX = X * CosVal - Y * SinVal
  nY = Y * CosVal + X * SinVal
End Sub
Public Sub FastRotateS2DXYo(RotAng As Long, X As Double, Y As Double, ox As Double, oy As Double, nX As Double, nY As Double)
Dim SinVal As Double, CosVal As Double
  RotAng = RotAng Mod 360
  SinVal = SinTable(RotAng)
  CosVal = CosTable(RotAng)
  X = X - ox
  Y = Y - oy
  nX = (X * CosVal - Y * SinVal) + ox
  nY = (Y * CosVal + X * SinVal) + oy
End Sub
Public Function FastRotate2DP(RotAng As Long, Point As TPoint2D) As TPoint2D
  Call FastRotateS2DXY(RotAng, Point.X, Point.Y, FastRotate2DP.X, FastRotate2DP.Y)
End Function
Public Function FastRotate2DPo(RotAng As Long, Point As TPoint2D, OPoint As TPoint2D) As TPoint2D
  Call FastRotateS2DXYo(RotAng, Point.X, Point.Y, OPoint.X, OPoint.Y, FastRotate2DPo.X, FastRotate2DPo.Y)
End Function
Public Function FastRotate2DS(RotAng As Long, Segment As TSegment2D) As TSegment2D
  With FastRotate2DS 'OK!!
    .P(0) = FastRotate2DP(RotAng, Segment.P(0))
    .P(1) = FastRotate2DP(RotAng, Segment.P(1))
  End With
End Function
Public Function FastRotate2DSo(RotAng As Long, Segment As TSegment2D, OPoint As TPoint2D) As TSegment2D
  With FastRotate2DSo 'OK!!
    .P(0) = FastRotate2DPo(RotAng, Segment.P(0), OPoint)
    .P(1) = FastRotate2DPo(RotAng, Segment.P(1), OPoint)
  End With
End Function
Public Function FastRotate2DT(RotAng As Long, Triangle As TTriangle2D) As TTriangle2D
  With FastRotate2DT 'OK!!
    .P(0) = FastRotate2DP(RotAng, Triangle.P(0))
    .P(1) = FastRotate2DP(RotAng, Triangle.P(1))
    .P(2) = FastRotate2DP(RotAng, Triangle.P(2))
  End With
End Function
Public Function FastRotate2DTo(RotAng As Long, Triangle As TTriangle2D, OPoint As TPoint2D) As TTriangle2D
  With FastRotate2DTo 'OK!!
    .P(0) = FastRotate2DPo(RotAng, Triangle.P(0), OPoint)
    .P(1) = FastRotate2DPo(RotAng, Triangle.P(1), OPoint)
    .P(2) = FastRotate2DPo(RotAng, Triangle.P(2), OPoint)
  End With
End Function
Public Function FastRotate2DQ(RotAng As Long, Quadix As TQuadix2D) As TQuadix2D
  With FastRotate2DQ 'OK!!
    .P(0) = FastRotate2DP(RotAng, Quadix.P(0))
    .P(1) = FastRotate2DP(RotAng, Quadix.P(1))
    .P(2) = FastRotate2DP(RotAng, Quadix.P(2))
  End With
End Function
Public Function FastRotate2DQo(RotAng As Long, Quadix As TQuadix2D, OPoint As TPoint2D) As TQuadix2D
  With FastRotate2DQo 'OK!!
    .P(0) = FastRotate2DPo(RotAng, Quadix.P(0), OPoint)
    .P(1) = FastRotate2DPo(RotAng, Quadix.P(1), OPoint)
    .P(2) = FastRotate2DPo(RotAng, Quadix.P(2), OPoint)
    .P(3) = FastRotate2DPo(RotAng, Quadix.P(3), OPoint)
  End With
End Function
Public Function FastRotate2DPg(RotAng As Long, Polygon As TPolygon2D) As TPolygon2D
Dim i As Long 'OK!!
  ReDim FastRotate2DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    FastRotate2DPg.Arr(i) = Rotate2DP(CDbl(RotAng), Polygon.Arr(i))
  Next
End Function
Public Function FastRotate2DPgo(RotAng As Long, Polygon As TPolygon2D, OPoint As TPoint2D) As TPolygon2D
Dim i As Long 'OK!!
  ReDim FastRotate2DPgo.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    FastRotate2DPgo.Arr(i) = Rotate2DPo(CDbl(RotAng), Polygon.Arr(i), OPoint)
  Next
End Function
Public Function FastRotate2DOo(RotAng As Long, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint2D:       FastRotate2DOo.Point2D = FastRotate2DP(RotAng, Obj.Point2D)
  Case geoSegment2D:   FastRotate2DOo.Segment2D = FastRotate2DS(RotAng, Obj.Segment2D)
  Case geoTriangle2D: FastRotate2DOo.Triangle2D = FastRotate2DT(RotAng, Obj.Triangle2D)
  Case geoQuadix2D:     FastRotate2DOo.Quadix2D = FastRotate2DQ(RotAng, Obj.Quadix2D)
  Case Else:            FastRotate2DOo = Obj
  End Select
End Function
Public Function FastRotate2DO(RotAng As Long, Obj As TGeometricObject, OPoint As TPoint2D) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint2D:       FastRotate2DO.Point2D = FastRotate2DPo(RotAng, Obj.Point2D, OPoint)
  Case geoSegment2D:   FastRotate2DO.Segment2D = FastRotate2DSo(RotAng, Obj.Segment2D, OPoint)
  Case geoTriangle2D: FastRotate2DO.Triangle2D = FastRotate2DTo(RotAng, Obj.Triangle2D, OPoint)
  Case geoQuadix2D:     FastRotate2DO.Quadix2D = FastRotate2DQo(RotAng, Obj.Quadix2D, OPoint)
  Case Else:            FastRotate2DO = Obj
  End Select
End Function
'(* Endof Fast Rotation 2D *)
Public Sub FastRotateS3DXY(Rx As Long, Ry As Long, Rz As Long, X As Double, Y As Double, z As Double, nX As Double, nY As Double, Nz As Double)
Dim TempX As Double, TempY As Double, TempZ As Double
Dim SinX As Double, SinY As Double, SinZ As Double
Dim CosX As Double, CosY As Double, CosZ As Double
  Rx = Rx Mod 360
  Ry = Ry Mod 360
  Rz = Rz Mod 360
  If Rx < 0 Then Rx = 360 + Rx
  If Ry < 0 Then Ry = 360 + Ry
  If Rz < 0 Then Rz = 360 + Rz
  SinX = SinTable(Rx)
  SinY = SinTable(Ry)
  SinZ = SinTable(Rz)
  CosX = CosTable(Rx)
  CosY = CosTable(Ry)
  CosZ = CosTable(Rz)
  TempY = Y * CosY - z * SinY
  TempZ = Y * SinY + z * CosY
  TempX = X * CosX - TempZ * SinX
  Nz = X * SinX + TempZ * CosX
  nX = TempX * CosZ - TempY * SinZ
  nY = TempX * SinZ + TempY * CosZ
End Sub
Public Sub FastRotateS3DXYo(Rx As Long, Ry As Long, Rz As Long, X As Double, Y As Double, z As Double, ox As Double, oy As Double, oz As Double, nX As Double, nY As Double, Nz As Double)
  Call FastRotateS3DXY(Rx, Ry, Rz, X - ox, Y - oy, z - oz, nX, nY, Nz)
  nX = nX + ox
  nY = nY + oy
  Nz = Nz + oz
End Sub
Public Function FastRotate3DP(Rx As Long, Ry As Long, Rz As Long, Point As TPoint3D) As TPoint3D
  Call FastRotateS3DXY(Rx, Ry, Rz, Point.X, Point.Y, Point.z, FastRotate3DP.X, FastRotate3DP.Y, FastRotate3DP.z)
End Function
Public Function FastRotate3DPo(Rx As Long, Ry As Long, Rz As Long, Point As TPoint3D, OPoint As TPoint3D) As TPoint3D
  Call FastRotateS3DXYo(Rx, Ry, Rz, Point.X, Point.Y, Point.z, OPoint.X, OPoint.Y, OPoint.z, FastRotate3DPo.X, FastRotate3DPo.Y, FastRotate3DPo.z)
End Function
Public Function FastRotate3DS(Rx As Long, Ry As Long, Rz As Long, Segment As TSegment3D) As TSegment3D
  With FastRotate3DS 'OK!!
    .P(0) = FastRotate3DP(Rx, Ry, Rz, Segment.P(0))
    .P(1) = FastRotate3DP(Rx, Ry, Rz, Segment.P(1))
  End With
End Function
Public Function FastRotate3DSo(Rx As Long, Ry As Long, Rz As Long, Segment As TSegment3D, OPoint As TPoint3D) As TSegment3D
  With FastRotate3DSo 'OK!!
    .P(0) = FastRotate3DPo(Rx, Ry, Rz, Segment.P(0), OPoint)
    .P(1) = FastRotate3DPo(Rx, Ry, Rz, Segment.P(1), OPoint)
  End With
End Function
Public Function FastRotate3DT(Rx As Long, Ry As Long, Rz As Long, Triangle As TTriangle3D) As TTriangle3D
  With FastRotate3DT 'OK!!
    .P(0) = FastRotate3DP(Rx, Ry, Rz, Triangle.P(0))
    .P(1) = FastRotate3DP(Rx, Ry, Rz, Triangle.P(1))
    .P(2) = FastRotate3DP(Rx, Ry, Rz, Triangle.P(2))
  End With
End Function
Public Function FastRotate3DTo(Rx As Long, Ry As Long, Rz As Long, Triangle As TTriangle3D, OPoint As TPoint3D) As TTriangle3D
  With FastRotate3DTo 'OK!!
    .P(0) = FastRotate3DPo(Rx, Ry, Rz, Triangle.P(0), OPoint)
    .P(1) = FastRotate3DPo(Rx, Ry, Rz, Triangle.P(1), OPoint)
    .P(2) = FastRotate3DPo(Rx, Ry, Rz, Triangle.P(2), OPoint)
  End With
End Function
Public Function FastRotate3DQ(Rx As Long, Ry As Long, Rz As Long, Quadix As TQuadix3D) As TQuadix3D
  With FastRotate3DQ 'OK!!
    .P(0) = FastRotate3DP(Rx, Ry, Rz, Quadix.P(0))
    .P(1) = FastRotate3DP(Rx, Ry, Rz, Quadix.P(1))
    .P(2) = FastRotate3DP(Rx, Ry, Rz, Quadix.P(2))
    .P(3) = FastRotate3DP(Rx, Ry, Rz, Quadix.P(3))
  End With
End Function
Public Function FastRotate3DQo(Rx As Long, Ry As Long, Rz As Long, Quadix As TQuadix3D, OPoint As TPoint3D) As TQuadix3D
  With FastRotate3DQo 'OK!!
    .P(0) = FastRotate3DPo(Rx, Ry, Rz, Quadix.P(0), OPoint)
    .P(1) = FastRotate3DPo(Rx, Ry, Rz, Quadix.P(1), OPoint)
    .P(2) = FastRotate3DPo(Rx, Ry, Rz, Quadix.P(2), OPoint)
    .P(3) = FastRotate3DPo(Rx, Ry, Rz, Quadix.P(3), OPoint)
  End With
End Function
Public Function FastRotate3DPg(Rx As Long, Ry As Long, Rz As Long, Polygon As TPolygon3D) As TPolygon3D
Dim i As Long 'OK!!
  ReDim FastRotate3DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    FastRotate3DPg.Arr(i) = FastRotate3DP(Rx, Ry, Rz, Polygon.Arr(i))
  Next
End Function
Public Function FastRotate3DPgo(Rx As Long, Ry As Long, Rz As Long, Polygon As TPolygon3D, OPoint As TPoint3D) As TPolygon3D
Dim i As Long 'OK!!
  ReDim FastRotate3DPgo.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    FastRotate3DPgo.Arr(i) = FastRotate3DPo(Rx, Ry, Rz, Polygon.Arr(i), OPoint)
  Next
End Function
Public Function FastRotate3DO(Rx As Long, Ry As Long, Rz As Long, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType 'OK!!
  Case geoPoint3D:       FastRotate3DO.Point3D = FastRotate3DP(Rx, Ry, Rz, Obj.Point3D)
  Case geoSegment3D:   FastRotate3DO.Segment3D = FastRotate3DS(Rx, Ry, Rz, Obj.Segment3D)
  Case geoTriangle3D: FastRotate3DO.Triangle3D = FastRotate3DT(Rx, Ry, Rz, Obj.Triangle3D)
  Case geoQuadix3D:     FastRotate3DO.Quadix3D = FastRotate3DQ(Rx, Ry, Rz, Obj.Quadix3D)
  Case Else:            FastRotate3DO = Obj
  End Select
End Function
Public Function FastRotate3DOo(Rx As Long, Ry As Long, Rz As Long, Obj As TGeometricObject, OPoint As TPoint3D) As TGeometricObject
  Select Case Obj.ObjectType 'OK!!
  Case geoPoint3D:       FastRotate3DOo.Point3D = FastRotate3DPo(Rx, Ry, Rz, Obj.Point3D, OPoint)
  Case geoSegment3D:   FastRotate3DOo.Segment3D = FastRotate3DSo(Rx, Ry, Rz, Obj.Segment3D, OPoint)
  Case geoTriangle3D: FastRotate3DOo.Triangle3D = FastRotate3DTo(Rx, Ry, Rz, Obj.Triangle3D, OPoint)
  Case geoQuadix3D:     FastRotate3DOo.Quadix3D = FastRotate3DQo(Rx, Ry, Rz, Obj.Quadix3D, OPoint)
  Case Else:            FastRotate3DOo = Obj
  End Select
End Function
'(* Endof Fast Rotation*)

'Translate
Public Function Translate2DP(Dx As Double, Dy As Double, Point As TPoint2D) As TPoint2D
  With Translate2DP
    .X = Point.X + Dx
    .Y = Point.Y + Dy
  End With
End Function
Public Function Translate2DL(Dx As Double, Dy As Double, Line As TLine2D) As TLine2D
  With Translate2DL 'OK!!
    .P(0).X = Line.P(0).X + Dx
    .P(0).Y = Line.P(0).Y + Dy
    .P(1).X = Line.P(1).X + Dx
    .P(1).Y = Line.P(1).Y + Dy
  End With
End Function
Public Function Translate2DS(Dx As Double, Dy As Double, Segment As TSegment2D) As TSegment2D
  With Translate2DS 'OK!!
    .P(0).X = Segment.P(0).X + Dx
    .P(0).Y = Segment.P(0).Y + Dy
    .P(1).X = Segment.P(1).X + Dx
    .P(1).Y = Segment.P(1).Y + Dy
  End With
End Function
Public Function Translate2DT(Dx As Double, Dy As Double, Triangle As TTriangle2D) As TTriangle2D
  With Translate2DT 'OK!!
    .P(0).X = Triangle.P(0).X + Dx
    .P(0).Y = Triangle.P(0).Y + Dy
    .P(1).X = Triangle.P(1).X + Dx
    .P(1).Y = Triangle.P(1).Y + Dy
    .P(2).X = Triangle.P(2).X + Dx
    .P(3).Y = Triangle.P(2).Y + Dy
  End With
End Function
Public Function Translate2DQ(Dx As Double, Dy As Double, Quadix As TQuadix2D) As TQuadix2D
  With Translate2DQ
    .P(0).X = Quadix.P(0).X + Dx
    .P(0).Y = Quadix.P(0).Y + Dy
    .P(1).X = Quadix.P(1).X + Dx
    .P(1).Y = Quadix.P(1).Y + Dy
    .P(2).X = Quadix.P(2).X + Dx
    .P(2).Y = Quadix.P(2).Y + Dy
    .P(3).X = Quadix.P(3).X + Dx
    .P(3).Y = Quadix.P(3).Y + Dy
  End With
End Function
Public Function Translate2DR(Dx As Double, Dy As Double, aRectangle As TRectangle) As TRectangle
  With Translate2DR
    .P(0).X = aRectangle.P(0).X + Dx
    .P(0).Y = aRectangle.P(0).Y + Dy
    .P(1).X = aRectangle.P(1).X + Dx
    .P(1).Y = aRectangle.P(1).Y + Dy
  End With
End Function
Public Function Translate2DCi(Dx As Double, Dy As Double, aCircle As TCircle) As TCircle
  With Translate2DCi
    .X = aCircle.X + Dx
    .Y = aCircle.Y + Dy
    .Radius = aCircle.Radius
  End With
End Function
Public Function Translate2DPg(Dx As Double, Dy As Double, Polygon As TPolygon2D) As TPolygon2D
Dim i As Long
  ReDim Translate2DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    Translate2DPg.Arr(i).X = Polygon.Arr(i).X + Dx
    Translate2DPg.Arr(i).Y = Polygon.Arr(i).Y + Dy
  Next
End Function
Public Function Translate2DPPg(Point As TPoint2D, Polygon As TPolygon2D) As TPolygon2D
  Translate2DPPg = Translate2DPg(Point.X, Point.Y, Polygon)
End Function
Public Function Translate2DO(Dx As Double, Dy As Double, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint2D:       Translate2DO.Point2D = Translate2DP(Dx, Dy, Obj.Point2D)
  Case geoSegment2D:   Translate2DO.Segment2D = Translate2DS(Dx, Dy, Obj.Segment2D)
  Case geoLine2D:         Translate2DO.Line2D = Translate2DL(Dx, Dy, Obj.Line2D)
  Case geoRectangle:   Translate2DO.Rectangle = Translate2DR(Dx, Dy, Obj.Rectangle)
  Case geoTriangle2D: Translate2DO.Triangle2D = Translate2DT(Dx, Dy, Obj.Triangle2D)
  Case geoQuadix2D:     Translate2DO.Quadix2D = Translate2DQ(Dx, Dy, Obj.Quadix2D)
  Case geoCircle:        Translate2DO.aCircle = Translate2DCi(Dx, Dy, Obj.aCircle)
  Case Else:          Translate2DO = Obj
  End Select
End Function
Public Function Translate2DdP(delta As Double, Point As TPoint2D) As TPoint2D
  Translate2DdP = Translate2DP(delta, delta, Point)
End Function
Public Function Translate2DdL(delta As Double, Line As TLine2D) As TLine2D
  Translate2DdL = Translate2DL(delta, delta, Line)
End Function
Public Function Translate2DdS(delta As Double, Segment As TSegment2D) As TSegment2D
  Translate2DdS = Translate2DS(delta, delta, Segment)
End Function
Public Function Translate2DdT(delta As Double, Triangle As TTriangle2D) As TTriangle2D
  Translate2DdT = Translate2DT(delta, delta, Triangle)
End Function
Public Function Translate2DdQ(delta As Double, Quadix As TQuadix2D) As TQuadix2D
  Translate2DdQ = Translate2DQ(delta, delta, Quadix)
End Function
Public Function Translate2DdR(delta As Double, aRectangle As TRectangle) As TRectangle
  Translate2DdR = Translate2DR(delta, delta, aRectangle)
End Function
Public Function Translate2DdCi(delta As Double, aCircle As TCircle) As TCircle
  Translate2DdCi = Translate2DCi(delta, delta, aCircle)
End Function
Public Function Translate2DdPg(delta As Double, Polygon As TPolygon2D) As TPolygon2D
  Translate2DdPg = Translate2DPg(delta, delta, Polygon)
End Function
Public Function Translate2DdO(delta As Double, Obj As TGeometricObject) As TGeometricObject
  Translate2DdO = Translate2DO(delta, delta, Obj)
End Function

Public Function Translate3DP(Dx As Double, Dy As Double, Dz As Double, Point As TPoint3D) As TPoint3D
  With Translate3DP
    .X = Point.X + Dx
    .Y = Point.Y + Dy
    .z = Point.z + Dz
  End With
End Function
Public Function Translate3DL(Dx As Double, Dy As Double, Dz As Double, Line As TLine3D) As TLine3D
  With Translate3DL 'OK!!
    .P(0).X = Line.P(0).X + Dx
    .P(0).Y = Line.P(0).Y + Dy
    .P(0).z = Line.P(0).z + Dz
    .P(1).X = Line.P(1).X + Dx
    .P(1).Y = Line.P(1).Y + Dy
    .P(1).z = Line.P(1).z + Dz
  End With
End Function
Public Function Translate3DS(Dx As Double, Dy As Double, Dz As Double, Segment As TSegment3D) As TSegment3D
  With Translate3DS 'OK!!
    .P(0).X = Segment.P(0).X + Dx
    .P(0).Y = Segment.P(0).Y + Dy
    .P(0).z = Segment.P(0).z + Dz
    .P(1).X = Segment.P(1).X + Dx
    .P(1).Y = Segment.P(1).Y + Dy
    .P(1).z = Segment.P(1).z + Dz
  End With
End Function
Public Function Translate3DT(Dx As Double, Dy As Double, Dz As Double, Triangle As TTriangle3D) As TTriangle3D
  With Translate3DT 'OK!!
    .P(0).X = Triangle.P(0).X + Dx
    .P(0).Y = Triangle.P(0).Y + Dy
    .P(0).z = Triangle.P(0).z + Dz
    .P(1).X = Triangle.P(1).X + Dx
    .P(1).Y = Triangle.P(1).Y + Dy
    .P(1).z = Triangle.P(1).z + Dz
    .P(2).X = Triangle.P(2).X + Dx
    .P(2).Y = Triangle.P(2).Y + Dy
    .P(2).z = Triangle.P(2).z + Dz
  End With
End Function
Public Function Translate3DQ(Dx As Double, Dy As Double, Dz As Double, Quadix As TQuadix3D) As TQuadix3D
  With Translate3DQ 'OK!!
    .P(0).X = Quadix.P(0).X + Dx
    .P(0).Y = Quadix.P(0).Y + Dy
    .P(0).z = Quadix.P(0).z + Dz
    .P(1).X = Quadix.P(1).X + Dx
    .P(1).Y = Quadix.P(1).Y + Dy
    .P(1).z = Quadix.P(1).z + Dz
    .P(2).X = Quadix.P(2).X + Dx
    .P(2).Y = Quadix.P(2).Y + Dy
    .P(2).z = Quadix.P(2).z + Dz
    .P(3).X = Quadix.P(3).X + Dx
    .P(3).Y = Quadix.P(3).Y + Dy
    .P(3).z = Quadix.P(3).z + Dz
  End With
End Function
Public Function Translate3DSph(Dx As Double, Dy As Double, Dz As Double, Sphere As TSphere) As TSphere
  With Translate3DSph
    .X = Sphere.X + Dx
    .Y = Sphere.Y + Dy
    .z = Sphere.z + Dz
  End With
End Function
Public Function Translate3DPg(Dx As Double, Dy As Double, Dz As Double, Polygon As TPolygon3D) As TPolygon3D
Dim i As Long
  ReDim Translate3DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    Translate3DPg.Arr(i).X = Polygon.Arr(i).X + Dx
    Translate3DPg.Arr(i).Y = Polygon.Arr(i).Y + Dy
    Translate3DPg.Arr(i).z = Polygon.Arr(i).z + Dz
  Next
End Function
Public Function Translate3DPPg(Point As TPoint3D, Polygon As TPolygon3D) As TPolygon3D
  Translate3DPPg = Translate3DPg(Point.X, Point.Y, Point.z, Polygon)
End Function
Public Function Translate3DO(Dx As Double, Dy As Double, Dz As Double, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint3D:       Translate3DO.Point3D = Translate3DP(Dx, Dy, Dz, Obj.Point3D)
  Case geoSegment3D:   Translate3DO.Segment3D = Translate3DS(Dx, Dy, Dz, Obj.Segment3D)
  Case geoLine3D:         Translate3DO.Line3D = Translate3DL(Dx, Dy, Dz, Obj.Line3D)
  Case geoTriangle3D: Translate3DO.Triangle3D = Translate3DT(Dx, Dy, Dz, Obj.Triangle3D)
  Case geoQuadix3D:     Translate3DO.Quadix3D = Translate3DQ(Dx, Dy, Dz, Obj.Quadix3D)
  Case geoSphere:         Translate3DO.Sphere = Translate3DSph(Dx, Dy, Dz, Obj.Sphere)
  Case Else:                     Translate3DO = Obj
  End Select
End Function
'(* Endof Translate *)

'Scale
Public Function Scale2DP(Dx As Double, Dy As Double, Point As TPoint2D) As TPoint2D
  With Scale2DP
    .X = Point.X * Dx
    .Y = Point.Y * Dy
  End With
End Function
Public Function Scale2DL(Dx As Double, Dy As Double, Line As TLine2D) As TLine2D
  With Scale2DL 'OK!!
    .P(0).X = Line.P(0).X * Dx
    .P(0).Y = Line.P(0).Y * Dy
    .P(1).X = Line.P(1).X * Dx
    .P(1).Y = Line.P(1).Y * Dy
  End With
End Function
Public Function Scale2DS(Dx As Double, Dy As Double, Segment As TSegment2D) As TSegment2D
  With Scale2DS 'OK!!
    .P(0).X = Segment.P(0).X * Dx
    .P(0).Y = Segment.P(0).Y * Dy
    .P(1).X = Segment.P(1).X * Dx
    .P(1).Y = Segment.P(1).Y * Dy
  End With
End Function
Public Function Scale2DT(Dx As Double, Dy As Double, Triangle As TTriangle2D) As TTriangle2D
  With Scale2DT 'OK!!
    .P(0).X = Triangle.P(0).X * Dx
    .P(0).Y = Triangle.P(0).Y * Dy
    .P(1).X = Triangle.P(1).X * Dx
    .P(1).Y = Triangle.P(1).Y * Dy
    .P(2).X = Triangle.P(2).X * Dx
    .P(2).Y = Triangle.P(2).Y * Dy
  End With
End Function
Public Function Scale2DQ(Dx As Double, Dy As Double, Quadix As TQuadix2D) As TQuadix2D
  With Scale2DQ 'OK!!
    .P(0).X = Quadix.P(0).X * Dx
    .P(0).Y = Quadix.P(0).Y * Dy
    .P(1).X = Quadix.P(1).X * Dx
    .P(1).Y = Quadix.P(1).Y * Dy
    .P(2).X = Quadix.P(2).X * Dx
    .P(2).Y = Quadix.P(2).Y * Dy
    .P(3).X = Quadix.P(3).X * Dx
    .P(3).Y = Quadix.P(3).Y * Dy
  End With
End Function
Public Function Scale2DR(Dx As Double, Dy As Double, aRectangle As TRectangle) As TRectangle
  With Scale2DR 'OK!!
    .P(0).X = aRectangle.P(0).X * Dx
    .P(0).Y = aRectangle.P(0).Y * Dy
    .P(1).X = aRectangle.P(1).X * Dx
    .P(1).Y = aRectangle.P(1).Y * Dy
  End With
End Function
Public Function Scale2DCi(Dr As Double, aCircle As TCircle) As TCircle
  With Scale2DCi
    .X = aCircle.X
    .Y = aCircle.Y
    .Radius = aCircle.Radius * Dr
  End With
End Function
Public Function Scale2DPg(Dx As Double, Dy As Double, Polygon As TPolygon2D) As TPolygon2D
Dim i As Long
  ReDim Scale2DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    Scale2DPg.Arr(i).X = Polygon.Arr(i).X * Dx
    Scale2DPg.Arr(i).Y = Polygon.Arr(i).Y * Dy
  Next
End Function
Public Function Scale2DO(Dx As Double, Dy As Double, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint2D:       Scale2DO.Point2D = Scale2DP(Dx, Dy, Obj.Point2D)
  Case geoSegment2D:   Scale2DO.Segment2D = Scale2DS(Dx, Dy, Obj.Segment2D)
  Case geoLine2D:         Scale2DO.Line2D = Scale2DL(Dx, Dy, Obj.Line2D)
  Case geoTriangle2D: Scale2DO.Triangle2D = Scale2DT(Dx, Dy, Obj.Triangle2D)
  Case geoQuadix2D:     Scale2DO.Quadix2D = Scale2DQ(Dx, Dy, Obj.Quadix2D)
  Case Else:                     Scale2DO = Obj
  End Select
End Function

Public Function Scale3DP(Dx As Double, Dy As Double, Dz As Double, Point As TPoint3D) As TPoint3D
  With Scale3DP
    .X = Point.X * Dx
    .Y = Point.Y * Dy
    .z = Point.z * Dz
  End With
End Function
Public Function Scale3DL(Dx As Double, Dy As Double, Dz As Double, Line As TLine3D) As TLine3D
  With Scale3DL 'OK!!
    .P(0).X = Line.P(0).X * Dx
    .P(0).Y = Line.P(0).Y * Dy
    .P(0).z = Line.P(0).z * Dz
    .P(1).X = Line.P(1).X * Dx
    .P(1).Y = Line.P(1).Y * Dy
    .P(1).z = Line.P(1).z * Dz
  End With
End Function
Public Function Scale3DS(Dx As Double, Dy As Double, Dz As Double, Segment As TSegment3D) As TSegment3D
  With Scale3DS 'OK!!
    .P(0).X = Segment.P(0).X * Dx
    .P(0).Y = Segment.P(0).Y * Dy
    .P(0).z = Segment.P(0).z * Dz
    .P(1).X = Segment.P(1).X * Dx
    .P(1).Y = Segment.P(1).Y * Dy
    .P(1).z = Segment.P(1).z * Dz
  End With
End Function
Public Function Scale3DT(Dx As Double, Dy As Double, Dz As Double, Triangle As TTriangle3D) As TTriangle3D
  With Scale3DT 'OK!!
    .P(0).X = Triangle.P(0).X * Dx
    .P(0).Y = Triangle.P(0).Y * Dy
    .P(0).z = Triangle.P(0).z * Dz
    .P(1).X = Triangle.P(1).X * Dx
    .P(1).Y = Triangle.P(1).Y * Dy
    .P(1).z = Triangle.P(1).z * Dz
    .P(2).X = Triangle.P(2).X * Dx
    .P(2).Y = Triangle.P(2).Y * Dy
    .P(2).z = Triangle.P(2).z * Dz
  End With
End Function
Public Function Scale3DQ(Dx As Double, Dy As Double, Dz As Double, Quadix As TQuadix3D) As TQuadix3D
  With Scale3DQ 'OK!!
    .P(0).X = Quadix.P(0).X * Dx
    .P(0).Y = Quadix.P(0).Y * Dy
    .P(0).z = Quadix.P(0).z * Dz
    .P(1).X = Quadix.P(1).X * Dx
    .P(1).Y = Quadix.P(1).Y * Dy
    .P(1).z = Quadix.P(1).z * Dz
    .P(2).X = Quadix.P(2).X * Dx
    .P(2).Y = Quadix.P(2).Y * Dy
    .P(2).z = Quadix.P(2).z * Dz
    .P(3).X = Quadix.P(3).X * Dx
    .P(3).Y = Quadix.P(3).Y * Dy
    .P(3).z = Quadix.P(3).z * Dz
  End With
End Function
Public Function Scale3DSph(Dr As Double, Sphere As TSphere) As TSphere
  With Scale3DSph
    .X = Sphere.X
    .Y = Sphere.Y
    .z = Sphere.z
    .Radius = Sphere.Radius * Dr
  End With
End Function
Public Function Scale3DPg(Dx As Double, Dy As Double, Dz As Double, Polygon As TPolygon3D) As TPolygon3D
Dim i As Long
  ReDim Scale3DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    Scale3DPg.Arr(i).X = Polygon.Arr(i).X * Dx
    Scale3DPg.Arr(i).Y = Polygon.Arr(i).Y * Dy
  Next
End Function
Public Function Scale3DO(Dx As Double, Dy As Double, Dz As Double, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint3D:       Scale3DO.Point3D = Scale3DP(Dx, Dy, Dz, Obj.Point3D)
  Case geoSegment3D:   Scale3DO.Segment3D = Scale3DS(Dx, Dy, Dz, Obj.Segment3D)
  Case geoLine3D:         Scale3DO.Line3D = Scale3DL(Dx, Dy, Dz, Obj.Line3D)
  Case geoTriangle3D: Scale3DO.Triangle3D = Scale3DT(Dx, Dy, Dz, Obj.Triangle3D)
  Case geoQuadix3D:     Scale3DO.Quadix3D = Scale3DQ(Dx, Dy, Dz, Obj.Quadix3D)
  Case Else:                     Scale3DO = Obj
  End Select
End Function
'(* Endof Scale *)

'ShearXAxis
Public Sub ShearXAxisS2DXY(Shear As Double, X As Double, Y As Double, nX As Double, nY As Double)
  nX = X + Shear * Y
  nY = Y
End Sub
Public Function ShearXAxis2DP(Shear As Double, Point As TPoint2D) As TPoint2D
  Call ShearXAxisS2DXY(Shear, Point.X, Point.Y, ShearXAxis2DP.X, ShearXAxis2DP.Y)
End Function
Public Function ShearXAxis2DS(Shear As Double, Segment As TSegment2D) As TSegment2D
  With ShearXAxis2DS 'OK!!
    .P(0) = ShearXAxis2DP(Shear, Segment.P(0)) '1'1
    .P(1) = ShearXAxis2DP(Shear, Segment.P(1)) '2'2
  End With
End Function
Public Function ShearXAxis2DT(Shear As Double, Triangle As TTriangle2D) As TTriangle2D
  With ShearXAxis2DT 'OK!! OhOh also jetzt schlägts dreizehn Maybe a bug??
    .P(0) = ShearXAxis2DP(Shear, Triangle.P(0)) '1'1
    .P(1) = ShearXAxis2DP(Shear, Triangle.P(1)) '2'2
    .P(2) = ShearXAxis2DP(Shear, Triangle.P(2)) '3'2
  End With
End Function
Public Function ShearXAxis2DQ(Shear As Double, Quadix As TQuadix2D) As TQuadix2D
  With ShearXAxis2DQ 'OK!! OhOh also jetzt schlägts dreizehn Maybe a bug??
    .P(0) = ShearXAxis2DP(Shear, Quadix.P(0)) '1'1
    .P(1) = ShearXAxis2DP(Shear, Quadix.P(1)) '2'2
    .P(2) = ShearXAxis2DP(Shear, Quadix.P(2)) '3'2
    .P(3) = ShearXAxis2DP(Shear, Quadix.P(3)) '3'2
  End With
End Function
Public Function ShearXAxis2DPg(Shear As Double, Polygon As TPolygon2D) As TPolygon2D
Dim i As Long
  ReDim ShearXAxis2DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    ShearXAxis2DPg.Arr(i) = ShearXAxis2DP(Shear, Polygon.Arr(i))
  Next
End Function
Public Function ShearXAxis2DO(Shear As Double, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint2D:       ShearXAxis2DO.Point2D = ShearXAxis2DP(Shear, Obj.Point2D)
  Case geoSegment2D:   ShearXAxis2DO.Segment2D = ShearXAxis2DS(Shear, Obj.Segment2D)
  Case geoTriangle2D: ShearXAxis2DO.Triangle2D = ShearXAxis2DT(Shear, Obj.Triangle2D)
  Case geoQuadix2D:     ShearXAxis2DO.Quadix2D = ShearXAxis2DQ(Shear, Obj.Quadix2D)
  Case Else:                     ShearXAxis2DO = Obj
  End Select
End Function
'(* Endof Shear 2D Along X-Axis *)

'ShearYAxis
Public Sub ShearYAxisS2DXY(Shear As Double, X As Double, Y As Double, nX As Double, nY As Double)
  nX = X
  nY = X * Shear + Y
End Sub
Public Function ShearYAxis2DP(Shear As Double, Point As TPoint2D) As TPoint2D
  Call ShearYAxisS2DXY(Shear, Point.X, Point.Y, ShearYAxis2DP.X, ShearYAxis2DP.Y)
End Function
Public Function ShearYAxis2DS(Shear As Double, Segment As TSegment2D) As TSegment2D
  With ShearYAxis2DS 'OK!!
    .P(0) = ShearYAxis2DP(Shear, Segment.P(0))
    .P(1) = ShearYAxis2DP(Shear, Segment.P(1))
  End With
End Function
Public Function ShearYAxis2DT(Shear As Double, Triangle As TTriangle2D) As TTriangle2D
  With ShearYAxis2DT 'OK!!
    .P(0) = ShearYAxis2DP(Shear, Triangle.P(0))
    .P(1) = ShearYAxis2DP(Shear, Triangle.P(1))
    .P(2) = ShearYAxis2DP(Shear, Triangle.P(2))
  End With
End Function
Public Function ShearYAxis2DQ(Shear As Double, Quadix As TQuadix2D) As TQuadix2D
  With ShearYAxis2DQ 'OK!!
    .P(0) = ShearYAxis2DP(Shear, Quadix.P(0)) 'siehe auch ShearXAxis
    .P(1) = ShearYAxis2DP(Shear, Quadix.P(1))
    .P(2) = ShearYAxis2DP(Shear, Quadix.P(2))
    .P(3) = ShearYAxis2DP(Shear, Quadix.P(3))
  End With
End Function
Public Function ShearYAxis2DPg(Shear As Double, Polygon As TPolygon2D) As TPolygon2D
Dim i As Long
  ReDim ShearYAxis2DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(Polygon.Arr)
    ShearYAxis2DPg.Arr(i) = ShearYAxis2DP(Shear, Polygon.Arr(i))
  Next
End Function
Public Function ShearYAxis2DO(Shear As Double, Obj As TGeometricObject) As TGeometricObject
  Select Case Obj.ObjectType
  Case geoPoint2D:       ShearYAxis2DO.Point2D = ShearYAxis2DP(Shear, Obj.Point2D)
  Case geoSegment2D:   ShearYAxis2DO.Segment2D = ShearYAxis2DS(Shear, Obj.Segment2D)
  Case geoTriangle2D: ShearYAxis2DO.Triangle2D = ShearYAxis2DT(Shear, Obj.Triangle2D)
  Case geoQuadix2D:     ShearYAxis2DO.Quadix2D = ShearYAxis2DQ(Shear, Obj.Quadix2D)
  Case Else:                     ShearYAxis2DO = Obj
  End Select
End Function
'(* Endof Shear 2D Along Y-Axis *)

'CenterAtLocation
Public Function CenterAtLocation2DPXY(Point As TPoint2D, X As Double, Y As Double) As TPoint2D
  CenterAtLocation2DPXY.X = X
  CenterAtLocation2DPXY.Y = Y
End Function
Public Function CenterAtLocation2DSXY(Segment As TSegment2D, X As Double, Y As Double) As TSegment2D
Dim cX As Double, cY As Double
  Call SegmentMidPointS2DS(Segment, cX, cY)
  CenterAtLocation2DSXY = Translate2DS(X - cX, Y - cY, Segment)
End Function
Public Function CenterAtLocation2DTXY(Triangle As TTriangle2D, X As Double, Y As Double) As TTriangle2D
Dim cX As Double, cY As Double
  Call CentroidS2DT(Triangle, cX, cY)
  CenterAtLocation2DTXY = Translate2DT(X - cX, Y - cY, Triangle)
End Function
Public Function CenterAtLocation2DRXY(Rectangle As TRectangle, X As Double, Y As Double) As TRectangle
Dim cX As Double, cY As Double 'OK!!
  cX = Abs(Rectangle.P(0).X - Rectangle.P(1).X) * 0.5
  cY = Abs(Rectangle.P(0).Y - Rectangle.P(1).Y) * 0.5
  CenterAtLocation2DRXY = Translate2DR(X - cX, Y - cY, Rectangle)
End Function
Public Function CenterAtLocation2DQXY(Quadix As TQuadix2D, X As Double, Y As Double) As TQuadix2D
  '(* to Be Completed *)
End Function
Public Function CenterAtLocation2DCiXY(aCircle As TCircle, X As Double, Y As Double) As TCircle
  CenterAtLocation2DCiXY.X = X
  CenterAtLocation2DCiXY.Y = Y
End Function
Public Function CenterAtLocation2DPgXY(Polygon As TPolygon2D, X As Double, Y As Double) As TPolygon2D
Dim cX As Double, cY As Double
  Call CentroidS2DPg(Polygon, cX, cY)
  CenterAtLocation2DPgXY = Translate2DPg(X - cX, Y - cY, Polygon)
End Function
Public Function CenterAtLocation2DPP(Point As TPoint2D, CPoint As TPoint2D) As TPoint2D
  CenterAtLocation2DPP = CenterAtLocation2DPXY(Point, CPoint.X, Point.Y)
End Function
Public Function CenterAtLocation2DSP(Segment As TSegment2D, CPoint As TPoint2D) As TSegment2D
  CenterAtLocation2DSP = CenterAtLocation2DSXY(Segment, CPoint.X, CPoint.Y)
End Function
Public Function CenterAtLocation2DTP(Triangle As TTriangle2D, CPoint As TPoint2D) As TTriangle2D
  CenterAtLocation2DTP = CenterAtLocation2DTXY(Triangle, CPoint.X, CPoint.Y)
End Function
Public Function CenterAtLocation2DRP(Rectangle As TRectangle, CPoint As TPoint2D) As TRectangle
  CenterAtLocation2DRP = CenterAtLocation2DRXY(Rectangle, CPoint.X, CPoint.Y)
End Function
Public Function CenterAtLocation2DQP(Quadix As TQuadix2D, CPoint As TPoint2D) As TQuadix2D
  CenterAtLocation2DQP = CenterAtLocation2DQXY(Quadix, CPoint.X, CPoint.Y)
End Function
Public Function CenterAtLocation2DCiP(aCircle As TCircle, CPoint As TPoint2D) As TCircle
  CenterAtLocation2DCiP = CenterAtLocation2DCiXY(aCircle, CPoint.X, CPoint.Y)
End Function
Public Function CenterAtLocation2DPgP(Polygon As TPolygon2D, CPoint As TPoint2D) As TPolygon2D
  CenterAtLocation2DPgP = CenterAtLocation2DPgXY(Polygon, CPoint.X, CPoint.Y)
End Function
'(* End Of Center At Location *)

'AABB = "Axis Aligned Bounding Box"
Public Function AABB2DS(Segment As TSegment2D) As TRectangle 'OK!!
  If Segment.P(0).X < Segment.P(1).X Then
    AABB2DS.P(0).X = Segment.P(0).X
    AABB2DS.P(1).X = Segment.P(1).X
  Else
    AABB2DS.P(0).X = Segment.P(1).X
    AABB2DS.P(1).X = Segment.P(0).X
  End If
  If Segment.P(0).Y < Segment.P(1).Y Then
    AABB2DS.P(0).Y = Segment.P(0).Y
    AABB2DS.P(1).Y = Segment.P(1).Y
  Else
    AABB2DS.P(0).Y = Segment.P(1).Y
    AABB2DS.P(1).Y = Segment.P(0).Y
  End If
End Function
Public Function AABB2DT(Triangle As TTriangle2D) As TRectangle 'OK!!
Dim i As Long
  With AABB2DT
    .P(0).X = Triangle.P(0).X
    .P(0).Y = Triangle.P(0).Y
    .P(1).X = Triangle.P(0).X
    .P(1).Y = Triangle.P(0).Y
  End With
  For i = 1 To 2 'OK!
    If Triangle.P(i).X < AABB2DT.P(0).X Then
      AABB2DT.P(1).X = Triangle.P(i).X
    ElseIf Triangle.P(i).X > AABB2DT.P(1).X Then
      AABB2DT.P(1).X = Triangle.P(i).X
    End If
    If Triangle.P(i).Y < AABB2DT.P(0).Y Then
      AABB2DT.P(1).Y = Triangle.P(i).Y
    ElseIf Triangle.P(i).Y > AABB2DT.P(1).Y Then
      AABB2DT.P(1).Y = Triangle.P(i).Y
    End If
  Next
End Function
Public Function AABB2DR(Rectangle As TRectangle) As TRectangle 'OK!!
  With AABB2DR
    .P(0).X = MinD(Rectangle.P(0).X, Rectangle.P(1).X)
    .P(0).Y = MinD(Rectangle.P(0).Y, Rectangle.P(1).Y)
    .P(1).X = MaxD(Rectangle.P(0).X, Rectangle.P(1).X)
    .P(1).Y = MaxD(Rectangle.P(0).Y, Rectangle.P(1).Y)
  End With
End Function
Public Function AABB2DQ(Quadix As TQuadix2D) As TRectangle 'OK!!
Dim i  As Long
  With AABB2DQ
    .P(0).X = Quadix.P(0).X
    .P(0).Y = Quadix.P(0).Y
    .P(1).X = Quadix.P(0).X
    .P(1).Y = Quadix.P(0).Y
  End With
  For i = 1 To 3
    If Quadix.P(i).X < AABB2DQ.P(0).X Then
      AABB2DQ.P(0).X = Quadix.P(i).X
    ElseIf Quadix.P(i).X > AABB2DQ.P(1).X Then
      AABB2DQ.P(1).X = Quadix.P(i).X
    End If
    If Quadix.P(i).Y < AABB2DQ.P(0).Y Then
      AABB2DQ.P(0).Y = Quadix.P(i).Y
    ElseIf Quadix.P(i).Y > AABB2DQ.P(1).Y Then
      AABB2DQ.P(1).Y = Quadix.P(i).Y
    End If
  Next
End Function
Public Function AABB2DCi(aCircle As TCircle) As TRectangle 'OK!!
  With AABB2DCi
    .P(0).X = aCircle.X - aCircle.Radius
    .P(0).Y = aCircle.Y - aCircle.Radius
    .P(1).X = aCircle.X + aCircle.Radius
    .P(1).Y = aCircle.Y + aCircle.Radius
  End With
End Function
Public Function AABB2DPg(Polygon As TPolygon2D) As TRectangle 'OK!!
Dim i  As Long
  With AABB2DPg
    .P(0).X = Polygon.Arr(0).X
    .P(0).Y = Polygon.Arr(0).Y
    .P(1).X = Polygon.Arr(0).X
    .P(1).Y = Polygon.Arr(0).Y
  End With
  For i = 1 To UBound(Polygon.Arr)
    If Polygon.Arr(i).X < AABB2DPg.P(0).X Then
      AABB2DPg.P(0).X = Polygon.Arr(i).X
    ElseIf Polygon.Arr(i).X > AABB2DPg.P(1).X Then
      AABB2DPg.P(1).X = Polygon.Arr(i).X
    End If
    If Polygon.Arr(i).Y < AABB2DPg.P(0).Y Then
      AABB2DPg.P(0).Y = Polygon.Arr(i).Y
    ElseIf Polygon.Arr(i).Y > AABB2DPg.P(1).Y Then
      AABB2DPg.P(1).Y = Polygon.Arr(i).Y
    End If
  Next
End Function
Public Function AABB2DCu(Curve As TPoint2DArray) As TRectangle 'OK!!
Dim i  As Long
  With AABB2DCu
    .P(0).X = Curve.Arr(0).X
    .P(0).Y = Curve.Arr(0).Y
    .P(1).X = Curve.Arr(0).X
    .P(1).Y = Curve.Arr(0).Y
  End With
  For i = 1 To UBound(Curve.Arr)
    If Curve.Arr(i).X < AABB2DCu.P(0).X Then
      AABB2DCu.P(0).X = Curve.Arr(i).X
    ElseIf Curve.Arr(i).X > AABB2DCu.P(1).X Then
      AABB2DCu.P(1).X = Curve.Arr(i).X
    End If
    If Curve.Arr(i).Y < AABB2DCu.P(0).Y Then
      AABB2DCu.P(0).Y = Curve.Arr(i).Y
    ElseIf Curve.Arr(i).Y > AABB2DCu.P(1).Y Then
      AABB2DCu.P(1).Y = Curve.Arr(i).Y
    End If
  Next
End Function
'Subs
Public Sub AABBS2DS(Segment As TSegment2D, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
Dim aRectangle As TRectangle 'OK!!
  aRectangle = AABB2DS(Segment)
  X1 = aRectangle.P(0).X
  Y1 = aRectangle.P(0).Y
  X2 = aRectangle.P(1).X
  Y2 = aRectangle.P(1).Y
End Sub
Public Sub AABBS2DT(Triangle As TTriangle2D, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
Dim aRectangle As TRectangle 'OK!!
  aRectangle = AABB2DT(Triangle)
  X1 = aRectangle.P(0).X
  Y1 = aRectangle.P(0).Y
  X2 = aRectangle.P(1).X
  Y2 = aRectangle.P(1).Y
End Sub
Public Sub AABBS2DR(Rectangle As TRectangle, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
Dim aRectangle As TRectangle 'OK!!
  aRectangle = AABB2DR(Rectangle)
  X1 = aRectangle.P(0).X
  Y1 = aRectangle.P(0).Y
  X2 = aRectangle.P(1).X
  Y2 = aRectangle.P(1).Y
End Sub
Public Sub AABBS2DQ(Quadix As TQuadix2D, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
Dim aRectangle As TRectangle 'OK!!
  aRectangle = AABB2DQ(Quadix)
  X1 = aRectangle.P(0).X
  Y1 = aRectangle.P(0).Y
  X2 = aRectangle.P(1).X
  Y2 = aRectangle.P(1).Y
End Sub
Public Sub AABBS2DCi(aCircle As TCircle, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
Dim aRectangle As TRectangle 'OK!!
  aRectangle = AABB2DCi(aCircle)
  X1 = aRectangle.P(0).X
  Y1 = aRectangle.P(0).Y
  X2 = aRectangle.P(1).X
  Y2 = aRectangle.P(1).Y
End Sub
Public Sub AABBS2DPg(Polygon As TPolygon2D, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
Dim aRectangle As TRectangle 'OK!!
  aRectangle = AABB2DPg(Polygon)
  X1 = aRectangle.P(0).X
  Y1 = aRectangle.P(0).Y
  X2 = aRectangle.P(1).X
  Y2 = aRectangle.P(1).Y
End Sub
Public Sub AABBS2DCu(Curve As TPoint2DArray, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
Dim aRectangle As TRectangle 'OK!!
  aRectangle = AABB2DCu(Curve)
  X1 = aRectangle.P(0).X
  Y1 = aRectangle.P(0).Y
  X2 = aRectangle.P(1).X
  Y2 = aRectangle.P(1).Y
End Sub
'(* End Of AABB *)

'ProjectPoint
Public Sub ProjectPointS2DSD(Srcx As Double, Srcy As Double, Dstx As Double, Dsty As Double, Dist As Double, nX As Double, nY As Double)
Dim DistRatio As Double
  DistRatio = Dist / Distance2DXY(Srcx, Srcy, Dstx, Dsty)
  nX = Srcx + DistRatio * (Dstx - Srcx)
  nY = Srcy + DistRatio * (Dsty - Srcy)
End Sub
Public Sub ProjectPointS3DSD(Srcx As Double, Srcy As Double, Srcz As Double, Dstx As Double, Dsty As Double, Dstz As Double, Dist As Double, nX As Double, nY As Double, Nz As Double)
Dim DistRatio As Double
  DistRatio = Dist / Distance3DXY(Srcx, Srcy, Srcz, Dstx, Dsty, Dstz)
  nX = Srcx + DistRatio * (Dstx - Srcx)
  nY = Srcy + DistRatio * (Dsty - Srcy)
  Nz = Srcz + DistRatio * (Dstz - Srcz)
End Sub
Public Sub ProjectPointS2DXY(Px As Double, Py As Double, Angle As Double, Dist As Double, nX As Double, nY As Double)
Dim Dx As Double, Dy As Double
  Dx = 0#
  Dy = 0#
  Select Case QuadrantA(Angle)
  Case 1: Dx = Cos(Angle * PIDiv180) * Dist
          Dy = Sin(Angle * PIDiv180) * Dist
  Case 2: Dx = Sin((Angle - 90#) * PIDiv180) * Dist * -1#
          Dy = Cos((Angle - 90#) * PIDiv180) * Dist
  Case 3: Dx = Cos((Angle - 180#) * PIDiv180) * Dist * -1#
          Dy = Sin((Angle - 180#) * PIDiv180) * Dist * -1#
  Case 4: Dx = Sin((Angle - 270#) * PIDiv180) * Dist
          Dy = Cos((Angle - 270#) * PIDiv180) * Dist * -1#
  End Select
  nX = Px + Dx
  nY = Py + Dy
End Sub
Public Function ProjectPoint2DSD(SrcPoint As TPoint2D, DstPoint As TPoint2D, Dist As Double) As TPoint2D
  Call ProjectPointS2DSD(SrcPoint.X, SrcPoint.Y, DstPoint.X, DstPoint.Y, Dist, ProjectPoint2DSD.X, ProjectPoint2DSD.Y)
End Function
Public Function ProjectPoint3DSD(SrcPoint As TPoint3D, DstPoint As TPoint3D, Dist As Double) As TPoint3D
  Call ProjectPointS3DSD(SrcPoint.X, SrcPoint.Y, SrcPoint.z, DstPoint.X, DstPoint.Y, DstPoint.z, Dist, ProjectPoint3DSD.X, ProjectPoint3DSD.Y, ProjectPoint3DSD.z)
End Function
Public Function ProjectPoint2DP(Point As TPoint2D, Angle As Double, Dist As Double) As TPoint2D
  Call ProjectPointS2DXY(Point.X, Point.Y, Angle, Dist, ProjectPoint2DP.X, ProjectPoint2DP.Y)
End Function
Public Sub ProjectPointS0(Px As Double, Py As Double, Dist As Double, nX As Double, nY As Double)
  nX = Px + Dist 'Dist = Distance
  nY = Py
End Sub
Public Sub ProjectPointS45(Px As Double, Py As Double, Dist As Double, nX As Double, nY As Double)
  nX = Px + 0.707106781186548 * Dist
  nY = Py + 0.707106781186548 * Dist
End Sub
Public Sub ProjectPointS90(Px As Double, Py As Double, Dist As Double, nX As Double, nY As Double)
  nX = Px
  nY = Py + Dist
End Sub
Public Sub ProjectPointS135(Px As Double, Py As Double, Dist As Double, nX As Double, nY As Double)
  nX = Px - 0.707106781186548 * Dist
  nY = Py + 0.707106781186548 * Dist
End Sub
Public Sub ProjectPointS180(Px As Double, Py As Double, Dist As Double, nX As Double, nY As Double)
  nX = Px - Dist
  nY = Py
End Sub
Public Sub ProjectPointS225(Px As Double, Py As Double, Dist As Double, nX As Double, nY As Double)
  nX = Px - 0.707106781186548 * Dist
  nY = Py - 0.707106781186548 * Dist
End Sub
Public Sub ProjectPointS270(Px As Double, Py As Double, Dist As Double, nX As Double, nY As Double)
  nX = Px
  nY = Py - Dist
End Sub
Public Sub ProjectPointS315(Px As Double, Py As Double, Dist As Double, nX As Double, nY As Double)
  nX = Px + 0.707106781186548 * Dist
  nY = Py - 0.707106781186548 * Dist
End Sub
Public Function ProjectPoint0(Point As TPoint2D, Dist As Double) As TPoint2D
  Call ProjectPointS0(Point.X, Point.Y, Dist, ProjectPoint0.X, ProjectPoint0.Y)
End Function
Public Function ProjectPoint45(Point As TPoint2D, Dist As Double) As TPoint2D
  Call ProjectPointS45(Point.X, Point.Y, Dist, ProjectPoint45.X, ProjectPoint45.Y)
End Function
Public Function ProjectPoint90(Point As TPoint2D, Dist As Double) As TPoint2D
  Call ProjectPointS90(Point.X, Point.Y, Dist, ProjectPoint90.X, ProjectPoint90.Y)
End Function
Public Function ProjectPoint135(Point As TPoint2D, Dist As Double) As TPoint2D
  Call ProjectPointS135(Point.X, Point.Y, Dist, ProjectPoint135.X, ProjectPoint135.Y)
End Function
Public Function ProjectPoint180(Point As TPoint2D, Dist As Double) As TPoint2D
  Call ProjectPointS180(Point.X, Point.Y, Dist, ProjectPoint180.X, ProjectPoint180.Y)
End Function
Public Function ProjectPoint225(Point As TPoint2D, Dist As Double) As TPoint2D
  Call ProjectPointS225(Point.X, Point.Y, Dist, ProjectPoint225.X, ProjectPoint225.Y)
End Function
Public Function ProjectPoint270(Point As TPoint2D, Dist As Double) As TPoint2D
  Call ProjectPointS270(Point.X, Point.Y, Dist, ProjectPoint270.X, ProjectPoint270.Y)
End Function
Public Function ProjectPoint315(Point As TPoint2D, Dist As Double) As TPoint2D
  Call ProjectPointS315(Point.X, Point.Y, Dist, ProjectPoint315.X, ProjectPoint315.Y)
End Function
'(* End of Project Point 2D *)


'ProjectObject
Public Function ProjectObject2DP(Point As TPoint2D, Angle As Double, Distance As Double) As TPoint2D
  ProjectObject2DP = ProjectPoint2DP(Point, Angle, Distance)
End Function
Public Function ProjectObject2DS(Segment As TSegment2D, Angle As Double, Distance As Double) As TSegment2D
  With ProjectObject2DS 'OK!!
    .P(0) = ProjectPoint2DP(Segment.P(0), Angle, Distance)
    .P(1) = ProjectPoint2DP(Segment.P(1), Angle, Distance)
  End With
End Function
Public Function ProjectObject2DT(Triangle As TTriangle2D, Angle As Double, Distance As Double) As TTriangle2D
  With ProjectObject2DT 'OK!!
    .P(0) = ProjectPoint2DP(Triangle.P(0), Angle, Distance)
    .P(1) = ProjectPoint2DP(Triangle.P(1), Angle, Distance)
    .P(2) = ProjectPoint2DP(Triangle.P(2), Angle, Distance)
  End With
End Function
Public Function ProjectObject2DQ(Quadix As TQuadix2D, Angle As Double, Distance As Double) As TQuadix2D
  With ProjectObject2DQ 'OK!!
    .P(0) = ProjectPoint2DP(Quadix.P(0), Angle, Distance)
    .P(1) = ProjectPoint2DP(Quadix.P(1), Angle, Distance)
    .P(2) = ProjectPoint2DP(Quadix.P(2), Angle, Distance)
    .P(3) = ProjectPoint2DP(Quadix.P(3), Angle, Distance)
  End With
End Function
Public Function ProjectObject2DCi(aCircle As TCircle, Angle As Double, Distance As Double) As TCircle
  Call ProjectPointS2DXY(aCircle.X, aCircle.Y, Angle, Distance, ProjectObject2DCi.X, ProjectObject2DCi.Y)
End Function
Public Function ProjectObject2DPg(Polygon As TPolygon2D, Angle As Double, Distance As Double) As TPolygon2D
Dim i As Long 'OK!!
  ReDim ProjectObject2DPg.Arr(UBound(Polygon.Arr))
  For i = 0 To UBound(ProjectObject2DPg.Arr)
    ProjectObject2DPg.Arr(i) = ProjectObject2DP(Polygon.Arr(i), Angle, Distance)
  Next
End Function
Public Function ProjectObject2DO(GeoObj As TGeometricObject, Angle As Double, Distance As Double) As TGeometricObject
  Select Case GeoObj.ObjectType 'OK!!
  Case geoPoint2D:       ProjectObject2DO.Point2D = ProjectObject2DP(GeoObj.Point2D, Angle, Distance)
  Case geoSegment2D:   ProjectObject2DO.Segment2D = ProjectObject2DS(GeoObj.Segment2D, Angle, Distance)
  Case geoTriangle2D: ProjectObject2DO.Triangle2D = ProjectObject2DT(GeoObj.Triangle2D, Angle, Distance)
  Case geoQuadix2D:     ProjectObject2DO.Quadix2D = ProjectObject2DQ(GeoObj.Quadix2D, Angle, Distance)
  Case geoCircle:        ProjectObject2DO.aCircle = ProjectObject2DCi(GeoObj.aCircle, Angle, Distance)
  Case Else:                     ProjectObject2DO = GeoObj
  End Select
End Function
'(* Endof Project Object *)

'CalculateBezierCoefficients
Public Sub CalculateBezierCoefficientsQ2D(Bezier As TQuadraticBezier2D, ax As Double, bx As Double, ay As Double, by As Double)
  bx = 2# * (Bezier.P(1).X - Bezier.P(0).X) 'OK!!
  by = 2# * (Bezier.P(1).Y - Bezier.P(0).Y)
  ax = Bezier.P(2).X - Bezier.P(0).X - bx
  ay = Bezier.P(2).Y - Bezier.P(0).Y - by
End Sub
Public Sub CalculateBezierCoefficientsQ3D(Bezier As TQuadraticBezier3D, ax As Double, bx As Double, ay As Double, by As Double, az As Double, bz As Double)
  bx = 2# * (Bezier.P(1).X - Bezier.P(0).X) 'OK!!
  by = 2# * (Bezier.P(1).Y - Bezier.P(0).Y)
  bz = 2# * (Bezier.P(1).z - Bezier.P(0).z)
  ax = Bezier.P(2).X - Bezier.P(0).X - bx
  ay = Bezier.P(2).Y - Bezier.P(0).Y - by
  az = Bezier.P(2).z - Bezier.P(0).z - bz
End Sub
Public Sub CalculateBezierCoefficientsC2D(Bezier As TCubicBezier2D, ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double)
  cX = 3# * (Bezier.P(1).X - Bezier.P(0).X) 'OK!!
  cY = 3# * (Bezier.P(1).Y - Bezier.P(0).Y)
  bx = 3# * (Bezier.P(2).X - Bezier.P(1).X) - cX
  by = 3# * (Bezier.P(2).Y - Bezier.P(1).Y) - cY
  ax = Bezier.P(3).X - Bezier.P(0).X - cX - bx
  ay = Bezier.P(3).Y - Bezier.P(0).Y - cY - by
End Sub
Public Sub CalculateBezierCoefficientsC3D(Bezier As TCubicBezier3D, ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double, az As Double, bz As Double, cz As Double)
  cX = 3# * (Bezier.P(1).X - Bezier.P(0).X) 'OK!!
  cY = 3# * (Bezier.P(1).Y - Bezier.P(0).Y)
  cz = 3# * (Bezier.P(1).z - Bezier.P(0).z)
  bx = 3# * (Bezier.P(2).X - Bezier.P(1).X) - cX
  by = 3# * (Bezier.P(2).Y - Bezier.P(1).Y) - cY
  bz = 3# * (Bezier.P(2).z - Bezier.P(1).z) - cz
  ax = Bezier.P(3).X - Bezier.P(0).X - cX - bx
  ay = Bezier.P(3).Y - Bezier.P(0).Y - cY - by
  az = Bezier.P(3).z - Bezier.P(0).z - cz - bz
End Sub
'(* Endof Calculate Bezier Values *)

'PointOnBezier
Public Function PointOnBezier2D(StartPoint As TPoint2D, ax As Double, bx As Double, ay As Double, by As Double, t As Double) As TPoint2D
Dim tSqr  As Double
  tSqr = t * t
  With PointOnBezier2D
    .X = (ax * tSqr) + (bx * t) + StartPoint.X
    .Y = (ay * tSqr) + (by * t) + StartPoint.Y
  End With
End Function
Public Function PointOnBezier3D(StartPoint As TPoint3D, ax As Double, bx As Double, ay As Double, by As Double, az As Double, bz As Double, t As Double) As TPoint3D
Dim tSqr  As Double
  tSqr = t * t
  With PointOnBezier3D
    .X = (ax * tSqr) + (bx * t) + StartPoint.X
    .Y = (ay * tSqr) + (by * t) + StartPoint.Y
    .z = (az * tSqr) + (bz * t) + StartPoint.z
  End With
End Function
Public Function PointOnBezier2Dc(StartPoint As TPoint2D, ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double, t As Double) As TPoint2D
Dim tSqr  As Double, tCube As Double
  tSqr = t * t
  tCube = tSqr * t
  With PointOnBezier2Dc
    .X = (ax * tCube) + (bx * tSqr) + (cX * t) + StartPoint.X
    .Y = (ay * tCube) + (by * tSqr) + (cY * t) + StartPoint.Y
  End With
End Function
Public Function PointOnBezier3Dc(StartPoint As TPoint3D, ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double, az As Double, bz As Double, cz As Double, t As Double) As TPoint3D
Dim tSqr  As Double, tCube As Double
  tSqr = t * t
  tCube = tSqr * t
  With PointOnBezier3Dc
    .X = (ax * tCube) + (bx * tSqr) + (cX * t) + StartPoint.X
    .Y = (ay * tCube) + (by * tSqr) + (cY * t) + StartPoint.Y
    .z = (az * tCube) + (bz * tSqr) + (cz * t) + StartPoint.z
  End With
End Function
'(* Endof Point On Bezier *)

'CreateBezier
Public Function CreateBezierQ2D(Bezier As TQuadraticBezier2D, PointCount As Long) As TPoint2DArray
Dim ax As Double, ay As Double, bx As Double, by As Double, dT As Double, t  As Double, i  As Long
  If PointCount = 0 Then Exit Function
  dT = 1# / (1# * PointCount - 1#)
  t = 0#
  ReDim CreateBezierQ2D.Arr(PointCount - 1)
  Call CalculateBezierCoefficientsQ2D(Bezier, ax, bx, ay, by)
  For i = 0 To PointCount - 1
    CreateBezierQ2D.Arr(i) = PointOnBezier2D(Bezier.P(0), ax, bx, ay, by, t)
    t = t + dT
  Next
End Function
Public Function CreateBezierQ3D(Bezier As TQuadraticBezier3D, PointCount As Long) As TPoint3DArray
Dim ax As Double, ay As Double, az As Double, bx As Double, by As Double, bz As Double, dT As Double, t  As Double, i  As Long
  If PointCount = 0 Then Exit Function
  dT = 1# / (1# * PointCount - 1#)
  t = 0#
  ReDim CreateBezierQ3D.Arr(PointCount - 1)
  Call CalculateBezierCoefficientsQ3D(Bezier, ax, bx, ay, by, az, bz)
  For i = 0 To PointCount - 1
    CreateBezierQ3D.Arr(i) = PointOnBezier3D(Bezier.P(0), ax, bx, ay, by, az, bz, t)
    t = t + dT
  Next
End Function
Public Function CreateBezierC2D(Bezier As TCubicBezier2D, PointCount As Long) As TPoint2DArray
Dim ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double
Dim dT As Double, t  As Double, i  As Long
  If PointCount = 0 Then Exit Function
  dT = 1# / (1# * PointCount - 1#)
  t = 0#
  ReDim CreateBezierC2D.Arr(PointCount - 1)
  Call CalculateBezierCoefficientsC2D(Bezier, ax, bx, cX, ay, by, cY)
  For i = 0 To PointCount - 1
    CreateBezierC2D.Arr(i) = PointOnBezier2Dc(Bezier.P(0), ax, bx, cX, ay, by, cY, t)
    t = t + dT
  Next
End Function
Public Function CreateBezierC3D(Bezier As TCubicBezier3D, PointCount As Long) As TPoint3DArray
Dim ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double, az As Double, bz As Double, cz As Double
Dim dT As Double, t  As Double, i  As Long
  If PointCount = 0 Then Exit Function
  dT = 1# / (1# * PointCount - 1#)
  t = 0#
  ReDim CreateBezierC3D.Arr(PointCount - 1)
  Call CalculateBezierCoefficientsC3D(Bezier, ax, bx, cX, ay, by, cY, az, bz, cz)
  For i = 0 To PointCount - 1
    CreateBezierC3D.Arr(i) = PointOnBezier3Dc(Bezier.P(0), ax, bx, cX, ay, by, cY, az, bz, cz, t)
    t = t + dT
  Next
End Function
'(* Endof Create Bezier *)

'CreateCurvePointBezier
Public Function CreateCurvePointBezierQ2D(Bezier As TQuadraticBezier2D, PointCount As Long) As TCurvePoint2DArray
Dim ax As Double, ay As Double, bx As Double, by As Double
Dim dT As Double, t  As Double, i  As Long
  If PointCount = 0 Then Exit Function
  dT = 1# / (1# * PointCount - 1#)
  t = 0#
  ReDim CreateCurvePointBezierQ2D.Arr(PointCount - 1)
  Call CalculateBezierCoefficientsQ2D(Bezier, ax, bx, ay, by)
  For i = 0 To PointCount - 1
    CreateCurvePointBezierQ2D.Arr(i) = EquateCurvePoint2DP(PointOnBezier2D(Bezier.P(0), ax, bx, ay, by, t), t)
    t = t + dT
  Next
End Function
Public Function CreateCurvePointBezierQ3D(Bezier As TQuadraticBezier3D, PointCount As Long) As TCurvePoint3DArray
Dim ax As Double, ay As Double, az As Double, bx As Double, by As Double, bz As Double
Dim dT As Double, t  As Double, i  As Long
  If PointCount = 0 Then Exit Function
  dT = 1# / (1# * PointCount - 1#)
  t = 0#
  ReDim CreateCurvePointBezierQ3D.Arr(PointCount - 1)
  Call CalculateBezierCoefficientsQ3D(Bezier, ax, bx, ay, by, az, bz)
  For i = 0 To PointCount - 1
    CreateCurvePointBezierQ3D.Arr(i) = EquateCurvePoint3DP(PointOnBezier3D(Bezier.P(0), ax, bx, ay, by, az, bz, t), t)
    t = t + dT
  Next
End Function
Public Function CreateCurvePointBezierC2D(Bezier As TCubicBezier2D, PointCount As Long) As TCurvePoint2DArray
Dim ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double
Dim dT As Double, t  As Double, i  As Long
  If PointCount = 0 Then Exit Function
  dT = 1# / (1# * PointCount - 1#)
  t = 0#
  ReDim CreateCurvePointBezierC2D.Arr(PointCount - 1)
  Call CalculateBezierCoefficientsC2D(Bezier, ax, bx, cX, ay, by, cY)
  For i = 0 To PointCount - 1
    CreateCurvePointBezierC2D.Arr(i) = EquateCurvePoint2DP(PointOnBezier2Dc(Bezier.P(0), ax, bx, cX, ay, by, cY, t), t)
    t = t + dT
  Next
End Function
Public Function CreateCurvePointBezierC3D(Bezier As TCubicBezier3D, PointCount As Long) As TCurvePoint3DArray
Dim ax As Double, bx As Double, cX As Double, ay As Double, by As Double, cY As Double, az As Double, bz As Double, cz As Double
Dim dT As Double, t  As Double, i  As Long
  If PointCount = 0 Then Exit Function
  dT = 1# / (1# * PointCount - 1#)
  t = 0#
  ReDim CreateCurvePointBezierC3D.Arr(PointCount - 1)
  Call CalculateBezierCoefficientsC3D(Bezier, ax, bx, cX, ay, by, cY, az, bz, cz)
  For i = 0 To PointCount - 1
    CreateCurvePointBezierC3D.Arr(i) = EquateCurvePoint3DP(PointOnBezier3Dc(Bezier.P(0), ax, bx, cX, ay, by, cY, az, bz, cz, t), t)
    t = t + dT
  Next
End Function
'(* Endof Create Curve Point Bezier *)

'CurveLength
Public Function CurveLengthQ2D(Bezier As TQuadraticBezier2D, PointCount As Long) As Double
Dim i As Long, Curve As TPoint2DArray
  CurveLengthQ2D = 0
  Curve = CreateBezierQ2D(Bezier, PointCount)
  For i = 0 To UBound(Curve.Arr) - 1 '- 2
    CurveLengthQ2D = CurveLengthQ2D + Distance2DP(Curve.Arr(i), Curve.Arr(i + 1))
  Next
  Erase Curve.Arr
End Function
Public Function CurveLengthQ3D(Bezier As TQuadraticBezier3D, PointCount As Long) As Double
Dim i As Long, Curve As TPoint3DArray
  CurveLengthQ3D = 0
  Curve = CreateBezierQ3D(Bezier, PointCount)
  For i = 0 To UBound(Curve.Arr) - 1 '- 2
    CurveLengthQ3D = CurveLengthQ3D + Distance3DP(Curve.Arr(i), Curve.Arr(i + 1))
  Next
  Erase Curve.Arr
End Function
Public Function CurveLengthC2D(Bezier As TCubicBezier2D, PointCount As Long) As Double
Dim i As Long, Curve As TPoint2DArray
  CurveLengthC2D = 0
  Curve = CreateBezierC2D(Bezier, PointCount)
  For i = 0 To UBound(Curve.Arr) - 1 '- 2
    CurveLengthC2D = CurveLengthC2D + Distance2DP(Curve.Arr(i), Curve.Arr(i + 1))
  Next
  Erase Curve.Arr
End Function
Public Function CurveLengthC3D(Bezier As TCubicBezier3D, PointCount As Long) As Double
Dim i As Long, Curve As TPoint3DArray
  CurveLengthC3D = 0
  Curve = CreateBezierC3D(Bezier, PointCount)
  For i = 0 To UBound(Curve.Arr) - 1 '- 2
    CurveLengthC3D = CurveLengthC3D + Distance3DP(Curve.Arr(i), Curve.Arr(i + 1))
  Next
  Erase Curve.Arr
End Function
'(* Endof CurveLength *)

'ShortenSegment
Public Sub ShortenSegmentS2D(Amount As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
Dim SegmentLength As Double, DistRatio As Double, Dx As Double, Dy As Double
  SegmentLength = Distance2DXY(X1, Y1, X2, Y2)
  If SegmentLength = 0 Then Exit Sub
  If SegmentLength < Amount Then
    Call SegmentMidPointS2DXY(X1, Y1, X2, Y2, X1, Y1)
    X2 = X1
    Y2 = Y1
    Exit Sub
  End If
  DistRatio = Amount / (2 * SegmentLength)
  Dx = X2 - X1
  Dy = Y2 - Y1
  X1 = X1 + DistRatio * Dx
  Y1 = Y1 + DistRatio * Dy
  X2 = X2 - DistRatio * Dx
  Y2 = Y2 - DistRatio * Dy
End Sub
Public Sub ShortenSegmentS3D(Amount As Double, X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double)
Dim SegmentLength As Double, DistRatio As Double, Dx As Double, Dy As Double, Dz As Double
  SegmentLength = Distance3DXY(X1, Y1, z1, X2, Y2, z2)
  If SegmentLength = 0 Then Exit Sub
  If SegmentLength <= Amount Then
    Call SegmentMidPointS3DXY(X1, Y1, z1, X2, Y2, z2, X1, Y1, z1)
    X2 = X1
    Y2 = Y1
    z2 = z1
    Exit Sub
  End If
  DistRatio = Amount / (2 * SegmentLength)
  Dx = X2 - X1
  Dy = Y2 - Y1
  Dz = z2 - z1
  X1 = X1 + DistRatio * Dx
  Y1 = Y1 + DistRatio * Dy
  z1 = z1 + DistRatio * Dz
  X2 = X2 - DistRatio * Dx
  Y2 = Y2 - DistRatio * Dy
  z2 = z2 - DistRatio * Dz
End Sub
Public Function ShortenSegment2D(Segment As TSegment2D, Amount As Double) As TSegment2D
  ShortenSegment2D = Segment
  Call ShortenSegmentS2D(Amount, ShortenSegment2D.P(0).X, ShortenSegment2D.P(0).Y, ShortenSegment2D.P(1).X, ShortenSegment2D.P(1).Y)
End Function
Public Function ShortenSegment3D(Segment As TSegment3D, Amount As Double) As TSegment3D
  ShortenSegment3D = Segment
  Call ShortenSegmentS3D(Amount, ShortenSegment3D.P(0).X, ShortenSegment3D.P(0).Y, ShortenSegment3D.P(0).z, ShortenSegment3D.P(1).X, ShortenSegment3D.P(1).Y, ShortenSegment3D.P(1).z)
End Function
'(* Endof ShortenSegment *)

'LengthenSegment
Public Sub LengthenSegmentS2D(Amount As Double, X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
Dim SegmentLength As Double, DistRatio As Double, Dx As Double, Dy As Double
  SegmentLength = Distance2DXY(X1, Y1, X2, Y2)
  If SegmentLength = 0 Then Exit Sub
  DistRatio = Amount / (2 * SegmentLength)
  Dx = X2 - X1
  Dy = Y2 - Y1
  X1 = X1 - DistRatio * Dx
  Y1 = Y1 - DistRatio * Dy
  X2 = X2 + DistRatio * Dx
  Y2 = Y2 + DistRatio * Dy
End Sub
Public Sub LengthenSegmentS3D(Amount As Double, X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double)
Dim SegmentLength As Double, DistRatio As Double, Dx As Double, Dy As Double, Dz As Double
  SegmentLength = Distance3DXY(X1, Y1, z1, X2, Y2, z2)
  If SegmentLength = 0 Then Exit Sub
  DistRatio = Amount / (2 * SegmentLength)
  Dx = X2 - X1
  Dy = Y2 - Y1
  Dz = z2 - z1
  X1 = X1 - DistRatio * Dx
  Y1 = Y1 - DistRatio * Dy
  z1 = z1 - DistRatio * Dz
  X2 = X2 + DistRatio * Dx
  Y2 = Y2 + DistRatio * Dy
  z2 = z2 + DistRatio * Dz
End Sub
Public Function LengthenSegment2D(Segment As TSegment2D, Amount As Double) As TSegment2D
  LengthenSegment2D = Segment
  Call LengthenSegmentS2D(Amount, LengthenSegment2D.P(0).X, LengthenSegment2D.P(0).Y, LengthenSegment2D.P(1).X, LengthenSegment2D.P(1).Y)
End Function
Public Function LengthenSegment3D(Segment As TSegment3D, Amount As Double) As TSegment3D
  LengthenSegment3D = Segment
  Call LengthenSegmentS3D(Amount, LengthenSegment3D.P(0).X, LengthenSegment3D.P(0).Y, LengthenSegment3D.P(0).z, LengthenSegment3D.P(1).X, LengthenSegment3D.P(1).Y, LengthenSegment3D.P(1).z)
End Function
'(* Endof Lengthen Segment *)

'EquatePoint
Public Function EquatePoint2D(X As Double, Y As Double) As TPoint2D
  With EquatePoint2D
    .X = X
    .Y = Y
  End With
End Function
Public Function EquatePoint3D(X As Double, Y As Double, z As Double) As TPoint3D
  With EquatePoint3D
    .X = X
    .Y = Y
    .z = z
  End With
End Function
Public Sub EquatePointS2D(X As Double, Y As Double, Point As TPoint2D)
  Point.X = X
  Point.Y = Y
End Sub
Public Sub EquatePointS3D(X As Double, Y As Double, z As Double, Point As TPoint3D)
  Point.X = X
  Point.Y = Y
  Point.z = z
End Sub
'(* End of Equate Point *)
'Public Function EquatePointPtr2D(x As Double, y As Double) As TPoint2DPtr
'  'New(Result)
'  Result = EquatePoint(x, y)
'End Function
'Public Function EquatePointPtr3D(x As Double, y As Double, z As Double) As TPoint3DPtr
'  'New(Result)
'  Result = EquatePoint(x, y, z)
'End Function
'Public Sub EquatePointPtrS2D(x As Double, y As Double, Point As TPoint2DPtr)
'  Point = EquatePointPtr(x, y)
'End Sub
'Public Sub EquatePointPtrS3D(x As Double, y As Double, z As Double, Point As TPoint3DPtr)
'  Point = EquatePointPtr(x, y, z)
'End Sub
''(* End of Equate Point Ptr*)

'EquateCurvePoint
Public Function EquateCurvePoint2DXY(X As Double, Y As Double, t As Double) As TCurvePoint2D
  With EquateCurvePoint2DXY
    .X = X
    .Y = Y
    .t = t
  End With
End Function
Public Function EquateCurvePoint3DXY(X As Double, Y As Double, z As Double, t As Double) As TCurvePoint3D
  With EquateCurvePoint3DXY
    .X = X
    .Y = Y
    .z = z
    .t = t
  End With
End Function
Public Function EquateCurvePoint2DP(Point As TPoint2D, t As Double) As TCurvePoint2D
  With EquateCurvePoint2DP
    .X = Point.X
    .Y = Point.Y
    .t = t
  End With
End Function
Public Function EquateCurvePoint3DP(Point As TPoint3D, t As Double) As TCurvePoint3D
  With EquateCurvePoint3DP
    .X = Point.X
    .Y = Point.Y
    .z = Point.z
    .t = t
  End With
End Function
'(* End of Equate Curve Point *)

'EquateSegment
Public Function EquateSegment2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As TSegment2D
  With EquateSegment2DXY
    .P(0).X = X1
    .P(0).Y = Y1
    .P(1).X = X2
    .P(1).Y = Y2
  End With
End Function
Public Function EquateSegment2DP(Point1 As TPoint2D, Point2 As TPoint2D) As TSegment2D
  EquateSegment2DP = EquateSegment2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Function EquateSegment3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As TSegment3D
  With EquateSegment3DXY
    .P(0).X = X1
    .P(0).Y = Y1
    .P(0).z = z1
    .P(1).X = X2
    .P(1).Y = Y2
    .P(1).z = z2
  End With
End Function
Public Function EquateSegment3DP(Point1 As TPoint3D, Point2 As TPoint3D) As TSegment3D
  EquateSegment3DP = EquateSegment3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z)
End Function
Public Sub EquateSegmentS2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, ByRef OutSegment As TSegment2D)
  With OutSegment
    .P(0).X = X1
    .P(0).Y = Y1
    .P(1).X = X2
    .P(1).Y = Y2
  End With
End Sub
Public Sub EquateSegmentS3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, ByRef OutSegment As TSegment3D)
  With OutSegment
    .P(0).X = X1
    .P(0).Y = Y1
    .P(0).z = z1
    .P(1).X = X2
    .P(1).Y = Y2
    .P(1).z = z2
  End With
End Sub
'(* End of Equate Segment *)

'EquateLine
Public Function EquateLine2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As TLine2D
  With EquateLine2DXY
    .P(0).X = X1: .P(0).Y = Y1
    .P(1).X = X2: .P(1).Y = Y2
  End With
End Function
Public Function EquateLine3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As TLine3D
  With EquateLine3DXY
    .P(0).X = X1: .P(0).Y = Y1: .P(0).z = z1
    .P(1).X = X2: .P(1).Y = Y2: .P(1).z = z2
  End With
End Function
Public Function EquateLine2DP(Point1 As TPoint2D, Point2 As TPoint2D) As TLine2D
  EquateLine2DP = EquateLine2DXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Function EquateLine3DP(Point1 As TPoint3D, Point2 As TPoint3D) As TLine3D
  EquateLine3DP = EquateLine3DXY(Point1.X, Point1.Y, Point1.z, Point1.X, Point2.Y, Point2.z)
End Function
Public Sub EquateLineS2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Line As TLine2D)
  With Line
    .P(0).X = X1: .P(0).Y = Y1
    .P(1).X = X2: .P(1).Y = Y2
  End With
End Sub
Public Sub EquateLineS3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, Line As TLine3D)
  With Line
    .P(0).X = X1: .P(0).Y = Y1: .P(0).z = z1
    .P(1).X = X2: .P(1).Y = Y2: .P(1).z = z2
  End With
End Sub
'(* End of Equate Line *)

'EquateQuadix
Public Function EquateQuadix2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As TQuadix2D
  With EquateQuadix2DXY
    .P(0).X = X1: .P(0).Y = Y1
    .P(1).X = X2: .P(1).Y = Y2
    .P(2).X = X3: .P(2).Y = Y3
    .P(3).X = X4: .P(3).Y = Y4
  End With
End Function
Public Function EquateQuadix3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double) As TQuadix3D
  With EquateQuadix3DXY
    .P(0).X = X1: .P(0).Y = Y1: .P(0).z = z1
    .P(1).X = X2: .P(1).Y = Y2: .P(1).z = z2
    .P(2).X = X3: .P(2).Y = Y3: .P(2).z = z3
    .P(3).X = X4: .P(3).Y = Y4: .P(3).z = z4
  End With
End Function
Public Function EquateQuadix2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Point4 As TPoint2D) As TQuadix2D
  EquateQuadix2DP = EquateQuadix2DXY(Point1.X, Point1.Y, _
                                     Point2.X, Point2.Y, _
                                     Point3.X, Point3.Y, _
                                     Point4.X, Point4.Y)
End Function
Public Function EquateQuadix3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Point4 As TPoint3D) As TQuadix3D
  EquateQuadix3DP = EquateQuadix3DXY(Point1.X, Point1.Y, Point1.z, _
                                     Point2.X, Point2.Y, Point1.z, _
                                     Point3.X, Point3.Y, Point1.z, _
                                     Point4.X, Point4.Y, Point1.z)
End Function
Public Sub EquateQuadixS2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, Quadix As TQuadix2D)
  With Quadix
    .P(0).X = X1: .P(0).Y = Y1
    .P(1).X = X2: .P(1).Y = Y2
    .P(2).X = X3: .P(2).Y = Y3
    .P(3).X = X4: .P(3).Y = Y4
  End With
End Sub
Public Sub EquateQuadixS3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double, Quadix As TQuadix3D)
  With Quadix
    .P(0).X = X1: .P(0).Y = Y1: .P(0).z = z1
    .P(1).X = X2: .P(1).Y = Y2: .P(1).z = z2
    .P(2).X = X3: .P(2).Y = Y3: .P(2).z = z3
    .P(3).X = X4: .P(3).Y = Y4: .P(3).z = z4
  End With
End Sub
'(* End of Equate Quadix *)

'EquateRectangle
Public Function EquateRectangleXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As TRectangle
  If X1 <= X2 Then
    EquateRectangleXY.P(0).X = X1
    EquateRectangleXY.P(1).X = X2
  Else
    EquateRectangleXY.P(0).X = X2
    EquateRectangleXY.P(1).X = X1
  End If
  If Y1 <= Y2 Then
    EquateRectangleXY.P(0).Y = Y1
    EquateRectangleXY.P(1).Y = Y2
  Else
    EquateRectangleXY.P(0).Y = Y2
    EquateRectangleXY.P(1).Y = Y1
  End If
End Function
Public Function EquateRectangleP(Point1 As TPoint2D, Point2 As TPoint2D) As TRectangle
  EquateRectangleP = EquateRectangleXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Sub EquateRectangleSXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Rect As TRectangle)
  If X1 <= X2 Then
    Rect.P(0).X = X1
    Rect.P(1).X = X2
  Else
    Rect.P(0).X = X2
    Rect.P(1).X = X1
  End If
  If Y1 <= Y2 Then
    Rect.P(0).Y = Y1
    Rect.P(1).Y = Y2
  Else
    Rect.P(0).Y = Y2
    Rect.P(1).Y = Y1
  End If
End Sub
Public Sub EquateRectangleSP(Point1 As TPoint2D, Point2 As TPoint2D, Rect As TRectangle)
  Call EquateRectangleSXY(Point1.X, Point1.Y, Point2.X, Point2.Y, Rect)
End Sub
'(* End of Equate Rectangle *)

'EquateCircle
Public Function EquateCircleXY(X As Double, Y As Double, r As Double) As TCircle
  EquateCircleXY.X = X
  EquateCircleXY.Y = Y
  EquateCircleXY.Radius = r
End Function
Public Function EquateCircleP(Point As TPoint2D, Radius As Double) As TCircle
  EquateCircleP = EquateCircleXY(Point.X, Point.Y, Radius)
End Function
Public Sub EquateCircleSXY(X As Double, Y As Double, r As Double, aCircle As TCircle)
  aCircle.X = X
  aCircle.Y = Y
  aCircle.Radius = r
End Sub
'(* End of Equate Circle *)

'EquateSphere
Public Function EquateSphereXY(X As Double, Y As Double, z As Double, r As Double) As TSphere
  EquateSphereXY.X = X
  EquateSphereXY.Y = Y
  EquateSphereXY.z = z
  EquateSphereXY.Radius = r
End Function
Public Sub EquateSphereSXY(X As Double, Y As Double, z As Double, r As Double, Sphere As TSphere)
  Sphere.X = X
  Sphere.Y = Y
  Sphere.z = z
  Sphere.Radius = r
End Sub
'(* End of Equate Sphere *)

'EquateTriangle
Public Function EquateTriangle2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As TTriangle2D
  With EquateTriangle2DXY
    .P(0).X = X1
    .P(0).Y = Y1
    .P(1).X = X2
    .P(1).Y = Y2
    .P(2).X = X3
    .P(2).Y = Y3
  End With
End Function
Public Function EquateTriangle3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As TTriangle3D
  With EquateTriangle3DXY
    .P(0).X = X1
    .P(1).X = X2
    .P(2).X = X3
    .P(0).Y = Y1
    .P(1).Y = Y2
    .P(2).Y = Y3
    .P(0).z = z1
    .P(1).z = z2
    .P(2).z = z3
  End With
End Function
Public Function EquateTriangle2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D) As TTriangle2D
  With EquateTriangle2DP
    .P(0) = Point1
    .P(1) = Point2
    .P(2) = Point3
  End With
End Function
Public Function EquateTriangle3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As TTriangle3D
  With EquateTriangle3DP
    .P(0) = Point1
    .P(1) = Point2
    .P(2) = Point3
  End With
End Function
Public Sub EquateTriangleS2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Triangle As TTriangle2D)
  Triangle.P(0).X = X1
  Triangle.P(1).X = X2
  Triangle.P(2).X = X3
  Triangle.P(0).Y = Y1
  Triangle.P(1).Y = Y2
  Triangle.P(2).Y = Y3
End Sub
Public Sub EquateTriangleS3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, Triangle As TTriangle3D)
  Triangle.P(0).X = X1
  Triangle.P(1).X = X2
  Triangle.P(2).X = X3
  Triangle.P(0).Y = Y1
  Triangle.P(1).Y = Y2
  Triangle.P(2).Y = Y3
  Triangle.P(0).z = z1
  Triangle.P(1).z = z2
  Triangle.P(2).z = z3
End Sub
Public Sub EquateTriangleS2DP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D, Triangle As TTriangle2D)
  Triangle.P(0) = Point1
  Triangle.P(1) = Point2
  Triangle.P(2) = Point3
End Sub
Public Sub EquateTriangleS3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Triangle As TTriangle3D)
  Triangle.P(0) = Point1
  Triangle.P(1) = Point2
  Triangle.P(2) = Point3
End Sub
'(* End of Equate Triangle *)

'EquatePlane
Public Function EquatePlane3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double) As TPlane2D
  With EquatePlane3DXY
    .a = Y1 * (z2 - z3) + Y2 * (z3 - z1) + Y3 * (z1 - z2)
    .b = z1 * (X2 - X3) + z2 * (X3 - X1) + z3 * (X1 - X2)
    .c = X1 * (Y2 - Y3) + X2 * (Y3 - Y1) + X3 * (Y1 - Y2)
    .D = -(X1 * (Y2 * z3 - Y3 * z2) + X2 * (Y3 * z1 - Y1 * z3) + X3 * (Y1 * z2 - Y2 * z1))
  End With
End Function
Public Sub EquatePlaneS3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, Plane As TPlane2D)
  Plane = EquatePlane3DXY(X1, Y1, z1, X2, Y2, z2, X3, Y3, z3)
End Sub
Public Function EquatePlane3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D) As TPlane2D
  EquatePlane3DP = EquatePlane3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z)
End Function
Public Sub EquatePlaneS3DP(Point1 As TPoint3D, Point2 As TPoint3D, Point3 As TPoint3D, Plane As TPlane2D)
  Call EquatePlaneS3DXY(Point1.X, Point1.Y, Point1.z, Point2.X, Point2.Y, Point2.z, Point3.X, Point3.Y, Point3.z, Plane)
End Sub
'(* End of Equate Plane *)

'EquateBezier
Public Sub EquateBezierS2D2XY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, Bezier As TQuadraticBezier2D)
  With Bezier
    .P(0).X = X1
    .P(0).Y = Y1
    .P(1).X = X2
    .P(1).Y = Y2
    .P(2).X = X3
    .P(2).Y = Y3
  End With
End Sub
Public Sub EquateBezierS3D2XY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, Bezier As TQuadraticBezier3D)
  With Bezier
    .P(0).X = X1
    .P(0).Y = Y1
    .P(0).z = z1
    .P(1).X = X2
    .P(1).Y = Y2
    .P(1).z = z2
    .P(2).X = X3
    .P(2).Y = Y3
    .P(2).z = z3
  End With
End Sub
Public Function EquateBezier2D2P(Pnt1 As TPoint2D, Pnt2 As TPoint2D, Pnt3 As TPoint2D) As TQuadraticBezier2D
  EquateBezier2D2P.P(0) = Pnt1
  EquateBezier2D2P.P(1) = Pnt2
  EquateBezier2D2P.P(2) = Pnt3
End Function
Public Function EquateBezier3D2P(Pnt1 As TPoint3D, Pnt2 As TPoint3D, Pnt3 As TPoint3D) As TQuadraticBezier3D
  EquateBezier3D2P.P(0) = Pnt1
  EquateBezier3D2P.P(1) = Pnt2
  EquateBezier3D2P.P(2) = Pnt3
End Function
Public Sub EquateBezierS2D3XY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, Bezier As TCubicBezier2D)
  Bezier.P(0).X = X1
  Bezier.P(0).Y = Y1
  Bezier.P(1).X = X2
  Bezier.P(1).Y = Y2
  Bezier.P(2).X = X3
  Bezier.P(2).Y = Y3
  Bezier.P(3).X = X4
  Bezier.P(3).Y = Y4
End Sub
Public Sub EquateBezierS3D3XY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double, X3 As Double, Y3 As Double, z3 As Double, X4 As Double, Y4 As Double, z4 As Double, Bezier As TCubicBezier3D)
  Bezier.P(0).X = X1
  Bezier.P(0).Y = Y1
  Bezier.P(0).z = z1
  Bezier.P(1).X = X2
  Bezier.P(1).Y = Y2
  Bezier.P(1).z = z2
  Bezier.P(2).X = X3
  Bezier.P(2).Y = Y3
  Bezier.P(2).z = z3
  Bezier.P(3).X = X4
  Bezier.P(3).Y = Y4
  Bezier.P(3).z = z4
End Sub
Public Function EquateBezier2D3P(Pnt1 As TPoint2D, Pnt2 As TPoint2D, Pnt3 As TPoint2D, Pnt4 As TPoint2D) As TCubicBezier2D
  EquateBezier2D3P.P(0) = Pnt1
  EquateBezier2D3P.P(1) = Pnt2
  EquateBezier2D3P.P(2) = Pnt3
  EquateBezier2D3P.P(3) = Pnt4
End Function
Public Function EquateBezier3D3P(Pnt1 As TPoint3D, Pnt2 As TPoint3D, Pnt3 As TPoint3D, Pnt4 As TPoint3D) As TCubicBezier3D
  EquateBezier3D3P.P(0) = Pnt1
  EquateBezier3D3P.P(1) = Pnt2
  EquateBezier3D3P.P(2) = Pnt3
  EquateBezier3D3P.P(3) = Pnt4
End Function
'(* End of Equate Bezier *)

'RectangleToQuadix
Public Function RectangleToQuadixXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As TQuadix2D
  With RectangleToQuadixXY
    .P(0) = EquatePoint2D(X1, Y1)
    .P(1) = EquatePoint2D(X2, Y1)
    .P(2) = EquatePoint2D(X2, Y2)
    .P(3) = EquatePoint2D(X1, Y2)
  End With
End Function
Public Function RectangleToQuadixP(Point1 As TPoint2D, Point2 As TPoint2D) As TQuadix2D
  RectangleToQuadixP = RectangleToQuadixXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Function
Public Function RectangleToQuadix(Rectangle As TRectangle) As TQuadix2D
  RectangleToQuadix = RectangleToQuadixP(Rectangle.P(0), Rectangle.P(1))
End Function
'(* Endof RectangleToQuadix *)

'TriangleToPolygon
Public Function TriangleToPolygonXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double) As TPolygon2D
  ReDim TriangleToPolygonXY.Arr(0 To 2) ' 3 - 1)
  With TriangleToPolygonXY
    .Arr(0) = EquatePoint2D(X1, Y1)
    .Arr(1) = EquatePoint2D(X2, Y2)
    .Arr(2) = EquatePoint2D(X3, Y3)
  End With
End Function
Public Function TriangleToPolygon(Triangle As TTriangle2D) As TPolygon2D
  TriangleToPolygon = TriangleToPolygonXY(Triangle.P(0).X, Triangle.P(0).Y, _
                                          Triangle.P(1).X, Triangle.P(1).Y, _
                                          Triangle.P(2).X, Triangle.P(2).Y)
End Function
'(* Endof TriangleToPolygon *)

'QuadixToPolygon
Public Function QuadixToPolygonXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double) As TPolygon2D
  ReDim QuadixToPolygonXY.Arr(0 To 3) '(4 - 1)
  QuadixToPolygonXY.Arr(0) = EquatePoint2D(X1, Y1)
  QuadixToPolygonXY.Arr(1) = EquatePoint2D(X2, Y2)
  QuadixToPolygonXY.Arr(2) = EquatePoint2D(X3, Y3)
  QuadixToPolygonXY.Arr(3) = EquatePoint2D(X4, Y4)
End Function
Public Function QuadixToPolygon(Quadix As TQuadix2D) As TPolygon2D
  QuadixToPolygon = QuadixToPolygonXY(Quadix.P(0).X, Quadix.P(0).Y, _
                                      Quadix.P(1).X, Quadix.P(1).Y, _
                                      Quadix.P(2).X, Quadix.P(2).Y, _
                                      Quadix.P(3).X, Quadix.P(3).Y)
End Function
'(* Endof QuadixToPolygon *)

'CircleToPolygon
Public Function CircleToPolygonXY(cX As Double, cY As Double, Radius As Double, PointCount As Long) As TPolygon2D
Dim i As Long, Angle As Double
  ReDim CircleToPolygonXY.Arr(PointCount - 1)
  Angle = 360# / (1# * PointCount)
  For i = 0 To PointCount - 1
    Call RotateS2DXYo(Angle * i, cX + Radius, cY, cX, cY, CircleToPolygonXY.Arr(i).X, CircleToPolygonXY.Arr(i).Y)
  Next
End Function
Public Function CircleToPolygon(aCircle As TCircle, PointCount As Long) As TPolygon2D
  CircleToPolygon = CircleToPolygonXY(aCircle.X, aCircle.Y, aCircle.Radius, PointCount)
End Function
'(* Endof Circle To Polygon *)

'SetGeometricObject
Public Sub SetGeometricObject2DP(Primitive As TPoint2D, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoPoint2D
  GeoObj.Point2D = Primitive
End Sub
Public Sub SetGeometricObject3DP(Primitive As TPoint3D, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoPoint3D
  GeoObj.Point3D = Primitive
End Sub
Public Sub SetGeometricObject2DL(Primitive As TLine2D, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoLine2D
  GeoObj.Line2D = Primitive
End Sub
Public Sub SetGeometricObject3DL(Primitive As TLine3D, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoLine3D
  GeoObj.Line3D = Primitive
End Sub
Public Sub SetGeometricObject2DS(Primitive As TSegment2D, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoSegment2D
  GeoObj.Segment2D = Primitive
End Sub
Public Sub SetGeometricObject3DS(Primitive As TSegment3D, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoSegment3D
  GeoObj.Segment3D = Primitive
End Sub
Public Sub SetGeometricObject2DT(Primitive As TTriangle2D, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoTriangle2D
  GeoObj.Triangle2D = Primitive
End Sub
Public Sub SetGeometricObject3DT(Primitive As TTriangle3D, GeoObj As TGeometricObject)
 GeoObj.ObjectType = geoTriangle3D
 GeoObj.Triangle3D = Primitive
End Sub
Public Sub SetGeometricObject2DQ(Primitive As TQuadix2D, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoQuadix2D
  GeoObj.Quadix2D = Primitive
End Sub
Public Sub SetGeometricObject3DQ(Primitive As TQuadix3D, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoQuadix3D
  GeoObj.Quadix3D = Primitive
End Sub
Public Sub SetGeometricObject2DR(Primitive As TRectangle, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoRectangle
  GeoObj.Rectangle = Primitive
End Sub
Public Sub SetGeometricObject2DCi(Primitive As TCircle, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoCircle
  GeoObj.aCircle = Primitive
End Sub
Public Sub SetGeometricObject3DSph(Primitive As TSphere, GeoObj As TGeometricObject)
  GeoObj.ObjectType = geoSphere
  GeoObj.Sphere = Primitive
End Sub
'(* Endof Set Geometric Object *)

Public Sub GenerateRandomPolygon(Bx1 As Double, By1 As Double, Bx2 As Double, By2 As Double, Polygon As TPolygon2D)
Dim i  As Long, Dx As Double, Dy As Double
  'Randomize
  Dx = Abs(Bx2 - Bx1)
  Dy = Abs(By2 - By1)
  'it could also be selfintersecting
  For i = 0 To UBound(Polygon.Arr)
    Polygon.Arr(i).X = Bx1 + GenerateRandomValue(Dx)
    Polygon.Arr(i).Y = By1 + GenerateRandomValue(Dy)
  Next
End Sub
'(* Endof GenerateRandomPolygon *)

'GenerateRandomPoints
Public Sub GenerateRandomPoints2DXY(Bx1 As Double, By1 As Double, Bx2 As Double, By2 As Double, PointList() As TPoint2D)
Dim i  As Long, Dx As Double, Dy As Double
  'Randomize
  Dx = Abs(Bx2 - Bx1)
  Dy = Abs(By2 - By1)
  For i = 0 To UBound(PointList)
    PointList(i).X = Bx1 + GenerateRandomValue(Dx)
    PointList(i).Y = By1 + GenerateRandomValue(Dy)
  Next
End Sub
Public Sub GenerateRandomPoints2DR(Rectangle As TRectangle, PointList() As TPoint2D)
  Call GenerateRandomPoints2DXY(Rectangle.P(0).X, Rectangle.P(0).Y, Rectangle.P(1).X, Rectangle.P(1).Y, PointList())
End Sub
Public Sub GenerateRandomPoints2DS(Segment As TSegment2D, PointList() As TPoint2D)
Dim i As Long, Dist As Double, Ratio As Double
  Dist = Distance2DS(Segment)
  If Dist = 0 Then Exit Sub
  For i = 0 To UBound(PointList)
    Ratio = GenerateRandomValue(Dist) / Dist
    PointList(i).X = Segment.P(0).X + Ratio * (Segment.P(1).X - Segment.P(0).X)
    PointList(i).Y = Segment.P(0).Y + Ratio * (Segment.P(1).Y - Segment.P(0).Y)
  Next
End Sub
Public Sub GenerateRandomPoints3DS(Segment As TSegment3D, PointList() As TPoint3D)
Dim i As Long, Dist As Double, Ratio As Double
  Dist = Distance3DS(Segment)
  If Dist = 0 Then Exit Sub
  For i = 0 To UBound(PointList)
    Ratio = GenerateRandomValue(Dist) / Dist
    PointList(i).X = Segment.P(0).X + Ratio * (Segment.P(1).X - Segment.P(0).X)
    PointList(i).Y = Segment.P(0).Y + Ratio * (Segment.P(1).Y - Segment.P(0).Y)
    PointList(i).z = Segment.P(0).z + Ratio * (Segment.P(1).z - Segment.P(0).z)
  Next
End Sub
Public Sub GenerateRandomPoints2DT(Triangle As TTriangle2D, PointList() As TPoint2D)
Dim a As Double, b As Double, c As Double, i As Long
  For i = 0 To UBound(PointList)
    a = RandomValue
    b = RandomValue
    If (a + b) > 1 Then
      a = 1 - a
      b = 1 - b
    End If
    c = (1 - a - b)
    PointList(i).X = (Triangle.P(0).X * a) + (Triangle.P(1).X * b) + (Triangle.P(2).X * c)
    PointList(i).Y = (Triangle.P(0).Y * a) + (Triangle.P(1).Y * b) + (Triangle.P(2).Y * c)
  Next
End Sub
Public Sub GenerateRandomPoints2DCi(aCircle As TCircle, PointList() As TPoint2D)
Dim i As Long, RandomAngle As Double, CPoint As TPoint2D
  CPoint = EquatePoint2D(aCircle.X, aCircle.Y)
  For i = 0 To UBound(PointList)
    RandomAngle = GenerateRandomValue(360)
    PointList(i).X = aCircle.X + aCircle.Radius * Sqr(RandomValue)
    PointList(i).Y = aCircle.Y
    PointList(i) = Rotate2DPo(RandomAngle, PointList(i), CPoint)
  Next
End Sub
Public Sub GenerateRandomPoints2DQ(Quadix As TQuadix2D, PointList() As TPoint2D)
'(*  Note As  It is assumed the input Quadix is convex. *)
Dim i As Long, a As Double, b As Double, a1 As Double, b1 As Double, a2 As Double, b2 As Double
Dim r1 As Double, r2 As Double, r3 As Double, r4 As Double
  For i = 0 To UBound(PointList)
    a = (2 * RandomValue) - 1
    b = (2 * RandomValue) - 1
    a1 = 1 - a
    a2 = 1 + a
    b1 = 1 - b
    b2 = 1 + b
    r1 = a1 * b1
    r2 = a2 * b1
    r3 = a2 * b2
    r4 = a1 * b2
    PointList(i).X = ((r1 * Quadix.P(0).X) + (r2 * Quadix.P(1).X) + (r3 * Quadix.P(2).X) + (r4 * Quadix.P(3).X)) * 0.25
    PointList(i).Y = ((r1 * Quadix.P(0).Y) + (r2 * Quadix.P(1).Y) + (r3 * Quadix.P(2).Y) + (r4 * Quadix.P(3).Y)) * 0.25
  Next
End Sub
'(* Endof Generate Random Points *)

'GenerateRandomPointsOnConvexPentagon
Public Sub GenerateRandomPointsOnConvexPentagon(Pentagon As TPolygon2D, PointList() As TPoint2D)
Dim Triangle As TTriangle2D, Quadix As TQuadix2D
Dim QRndPoint() As TPoint2D, TRndPoint() As TPoint2D
Dim i As Long, QRndPntCnt As Long, TRndPntCnt As Long
tryE: On Error GoTo 0
  If UBound(Pentagon.Arr) <> 5 Then Exit Sub
  If Not ConvexPolygon(Pentagon) Then Exit Sub
  Quadix = EquateQuadix2DP(Pentagon.Arr(0), Pentagon.Arr(1), Pentagon.Arr(2), Pentagon.Arr(3))
  Triangle = EquateTriangle2DP(Pentagon.Arr(3), Pentagon.Arr(4), Pentagon.Arr(0))
  QRndPntCnt = Round(1# * (UBound(PointList) + 1) * (Area2DQ(Quadix) / Area2DPg(Pentagon)))
  TRndPntCnt = (UBound(PointList) + 1) - QRndPntCnt
  ReDim QRndPoint(QRndPntCnt - 1)
  ReDim TRndPoint(TRndPntCnt - 1)
  Call GenerateRandomPoints2DQ(Quadix, QRndPoint)
  Call GenerateRandomPoints2DT(Triangle, TRndPoint)
  For i = 0 To UBound(QRndPoint)
    PointList(i) = QRndPoint(i)
  Next
  For i = 0 To UBound(TRndPoint)
    PointList(i + UBound(QRndPoint)) = TRndPoint(i)
  Next
  Erase QRndPoint
  Erase TRndPoint
End Sub
'(* Endof Generate Random Points On Convex Pentagon *)

'GenerateRandomPointsOnConvexHexagon
Public Sub GenerateRandomPointsOnConvexHexagon(Hexagon As TPolygon2D, PointList() As TPoint2D)
Dim Quadix1 As TQuadix2D, Quadix2 As TQuadix2D
Dim Q1RndPoint() As TPoint2D, Q2RndPoint() As TPoint2D
Dim i As Long, Q1RndPntCnt As Long, Q2RndPntCnt As Long
  If UBound(Hexagon.Arr) <> 6 Then Exit Sub
  If Not ConvexPolygon(Hexagon) Then Exit Sub
  Quadix1 = EquateQuadix2DP(Hexagon.Arr(0), Hexagon.Arr(1), Hexagon.Arr(2), Hexagon.Arr(5))
  Quadix2 = EquateQuadix2DP(Hexagon.Arr(2), Hexagon.Arr(3), Hexagon.Arr(4), Hexagon.Arr(5))
  Q1RndPntCnt = CLng(1# * (UBound(PointList) + 1) * (Area2DQ(Quadix1) / Area2DPg(Hexagon)))
  Q2RndPntCnt = (UBound(PointList) + 1) - Q1RndPntCnt
  ReDim Q1RndPoint(Q1RndPntCnt - 1)
  ReDim Q2RndPoint(Q2RndPntCnt - 1)
  Call GenerateRandomPoints2DQ(Quadix1, Q1RndPoint)
  Call GenerateRandomPoints2DQ(Quadix2, Q2RndPoint)
  For i = 0 To UBound(Q1RndPoint)
    PointList(i) = Q1RndPoint(i)
  Next
  For i = 0 To UBound(Q2RndPoint)
    PointList(i + UBound(Q1RndPoint)) = Q2RndPoint(i)
  Next
  Erase Q1RndPoint
  Erase Q2RndPoint
End Sub
'(* Endof Generate Random Points On Convex Hexagon *)

'GenerateRandomPointsOnConvexHeptagon
Public Sub GenerateRandomPointsOnConvexHeptagon(Heptagon As TPolygon2D, PointList() As TPoint2D)
Dim Quadix1 As TQuadix2D, Quadix2 As TQuadix2D, Triangle As TTriangle2D
Dim Q1RndPoint() As TPoint2D, Q2RndPoint() As TPoint2D, TRndPoint() As TPoint2D
Dim i As Long, Q1RndPntCnt As Long, Q2RndPntCnt As Long, TRndPntCnt  As Long
  If UBound(Heptagon.Arr) <> 7 Then Exit Sub
  If Not ConvexPolygon(Heptagon) Then Exit Sub
  Quadix1 = EquateQuadix2DP(Heptagon.Arr(0), Heptagon.Arr(1), Heptagon.Arr(2), Heptagon.Arr(3))
  Quadix2 = EquateQuadix2DP(Heptagon.Arr(3), Heptagon.Arr(4), Heptagon.Arr(5), Heptagon.Arr(6))
  Triangle = EquateTriangle2DP(Heptagon.Arr(6), Heptagon.Arr(0), Heptagon.Arr(3))
  Q1RndPntCnt = Round(1# * (UBound(PointList) + 1) * (Area2DQ(Quadix1) / Area2DPg(Heptagon)))
  Q2RndPntCnt = Round(1# * ((UBound(PointList) + 1) - Q1RndPntCnt) * (Area2DQ(Quadix2) / (Area2DQ(Quadix2) + Area2DT(Triangle))))
  TRndPntCnt = (UBound(PointList) + 1) - Q1RndPntCnt - Q2RndPntCnt
  ReDim Q1RndPoint(0 To Q1RndPntCnt - 1)
  ReDim Q2RndPoint(0 To Q2RndPntCnt - 1)
  ReDim TRndPoint(0 To TRndPntCnt - 1)
  Call GenerateRandomPoints2DQ(Quadix1, Q1RndPoint)
  Call GenerateRandomPoints2DQ(Quadix2, Q2RndPoint)
  Call GenerateRandomPoints2DT(Triangle, TRndPoint)
  For i = 0 To UBound(Q1RndPoint)
    PointList(i) = Q1RndPoint(i)
  Next
  For i = 0 To UBound(Q2RndPoint)
    PointList(i + UBound(Q1RndPoint)) = Q2RndPoint(i)
  Next
  For i = 0 To UBound(TRndPoint)
    PointList(i + UBound(Q1RndPoint) + UBound(Q2RndPoint)) = TRndPoint(i)
  Next
  Erase Q1RndPoint
  Erase Q2RndPoint
  Erase TRndPoint
End Sub
'(* Endof Generate Random Points On Convex Heptagon *)

'GenerateRandomPointsOnConvexOctagon
Public Sub GenerateRandomPointsOnConvexOctagon(Octagon As TPolygon2D, PointList() As TPoint2D)
Dim Quadix1 As TQuadix2D, Quadix2 As TQuadix2D, Quadix3 As TQuadix2D
Dim Q1RndPoint() As TPoint2D, Q2RndPoint() As TPoint2D, Q3RndPoint() As TPoint2D
Dim i As Long, Q1RndPntCnt As Long, Q2RndPntCnt As Long, Q3RndPntCnt As Long
  If UBound(Octagon.Arr) <> 8 Then Exit Sub
  If Not ConvexPolygon(Octagon) Then Exit Sub
  Quadix1 = EquateQuadix2DP(Octagon.Arr(0), Octagon.Arr(1), Octagon.Arr(2), Octagon.Arr(3))
  Quadix2 = EquateQuadix2DP(Octagon.Arr(3), Octagon.Arr(4), Octagon.Arr(5), Octagon.Arr(6))
  Quadix3 = EquateQuadix2DP(Octagon.Arr(0), Octagon.Arr(3), Octagon.Arr(6), Octagon.Arr(7))
  Q1RndPntCnt = Round(1# * (UBound(PointList) + 1) * (Area2DQ(Quadix1) / Area2DPg(Octagon)))
  Q2RndPntCnt = Round(1# * ((UBound(PointList) + 1) - Q1RndPntCnt) * (Area2DQ(Quadix2) / (Area2DQ(Quadix2) + Area2DQ(Quadix3))))
  Q3RndPntCnt = (UBound(PointList) + 1) - Q1RndPntCnt - Q2RndPntCnt
  ReDim Q1RndPoint(0 To Q1RndPntCnt - 1)
  ReDim Q2RndPoint(0 To Q2RndPntCnt - 1)
  ReDim Q3RndPoint(0 To Q3RndPntCnt - 1)
  Call GenerateRandomPoints2DQ(Quadix1, Q1RndPoint)
  Call GenerateRandomPoints2DQ(Quadix2, Q2RndPoint)
  Call GenerateRandomPoints2DQ(Quadix3, Q3RndPoint)
  For i = 0 To UBound(Q1RndPoint)
    PointList(i) = Q1RndPoint(i)
  Next
  For i = 0 To UBound(Q2RndPoint)
    PointList(i + UBound(Q1RndPoint)) = Q2RndPoint(i)
  Next
  For i = 0 To UBound(Q2RndPoint)
    PointList(i + (UBound(Q1RndPoint) + 1) + (UBound(Q2RndPoint) + 1)) = Q3RndPoint(i)
  Next
  Erase Q1RndPoint
  Erase Q2RndPoint
  Erase Q3RndPoint
End Sub
'(* Endof Generate Random Points On Convex Octagon *)

'GenerateRandomTriangle
Public Sub GenerateRandomTriangle(Bx1 As Double, By1 As Double, Bx2 As Double, By2 As Double, Triangle As TTriangle2D)
Dim Dx As Double, Dy As Double
  Dx = Abs(Bx2 - Bx1)
  Dy = Abs(By2 - By1)
  Do
    With Triangle
      .P(0).X = Bx1 + GenerateRandomValue(Dx)
      .P(0).Y = By1 + GenerateRandomValue(Dy)
      .P(1).X = Bx1 + GenerateRandomValue(Dx)
      .P(1).Y = By1 + GenerateRandomValue(Dy)
      .P(2).X = Bx1 + GenerateRandomValue(Dx)
      .P(2).Y = By1 + GenerateRandomValue(Dy)
    End With
  Loop Until Not IsDegenerate2DT(Triangle)
End Sub
'(* Endof Generate Random Triangle *)

'GenerateRandomQuadix
Public Sub GenerateRandomQuadix(Bx1 As Double, By1 As Double, Bx2 As Double, By2 As Double, Quadix As TQuadix2D)
'Dim Dx As Long, Dy As Long
Dim Dx As Double, Dy As Double
  'Dx = CLng(Abs(Bx2 - Bx1) - 1#)
  'Dy = CLng(Abs(By2 - By1) - 1#)
  Dx = Abs(Bx2 - Bx1)
  Dy = Abs(By2 - By1)
  Do
    With Quadix
      .P(0).X = Bx1 + GenerateRandomValue(Dx) 'RandomL(Dx) + RandomValue
      .P(0).Y = By1 + GenerateRandomValue(Dy) 'RandomL(Dy) + RandomValue
      .P(1).X = Bx1 + GenerateRandomValue(Dx) 'RandomL(Dx) + RandomValue
      .P(1).Y = By1 + GenerateRandomValue(Dy) 'RandomL(Dy) + RandomValue
      .P(2).X = Bx1 + GenerateRandomValue(Dx) 'RandomL(Dx) + RandomValue
      .P(2).Y = By1 + GenerateRandomValue(Dy) 'RandomL(Dy) + RandomValue
      .P(3).X = Bx1 + GenerateRandomValue(Dx) 'RandomL(Dx) + RandomValue
      .P(3).Y = By1 + GenerateRandomValue(Dy) 'RandomL(Dy) + RandomValue
    End With
  Loop Until (Not IsDegenerate2DQ(Quadix)) And ConvexQuadix(Quadix)
End Sub
'(* Endof Generate Random Quadix *)

'GenerateRandomRectangle
Public Sub GenerateRandomRectangle(Bx1 As Double, By1 As Double, Bx2 As Double, By2 As Double, aRectangle As TRectangle)
'Dim Dx As Long, Dy As Long
Dim Dx As Double, Dy As Double
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
Dim TmpXY As Double
  'Dx = CLng(Abs(Bx2 - Bx1)) '- 1#)
  'Dy = CLng(Abs(By2 - By1)) '- 1#)
  Dx = Abs(Bx2 - Bx1)
  Dy = Abs(By2 - By1)
  X1 = Bx1 + GenerateRandomValue(Dx) 'RandomL(Dx) '+ RandomValue
  Y1 = By1 + GenerateRandomValue(Dy) 'RandomL(Dy) '+ RandomValue
  X2 = Bx1 + GenerateRandomValue(Dx) 'RandomL(Dx) '+ RandomValue
  Y2 = By1 + GenerateRandomValue(Dy) 'RandomL(Dy) '+ RandomValue
  If X1 > X2 Then
    Call SwapS(X1, X2)
  End If
  If Y1 > Y2 Then
    Call SwapS(Y1, Y2)
  End If
  With aRectangle
    .P(0).X = X1
    .P(0).Y = Y1
    .P(1).X = X2
    .P(1).Y = Y2
  End With
End Sub
'(* Endof Generate Random Rectangle *)

'GenerateRandomCircle
Public Sub GenerateRandomCircle(Bx1 As Double, By1 As Double, Bx2 As Double, By2 As Double, aCircle As TCircle)
Dim Dx As Double, Dy As Double
  Dx = Abs(Bx2 - Bx1) - 1#
  Dy = Abs(By2 - By1) - 1#
  aCircle.Radius = RandomL(CLng(MinD(Dx, Dy) * 0.5)) + RandomValue
  aCircle.X = Bx1 + aCircle.Radius + RandomL(CLng(Dx - (2 * aCircle.Radius))) + RandomValue
  aCircle.Y = By1 + aCircle.Radius + RandomL(CLng(Dy - (2 * aCircle.Radius))) + RandomValue
End Sub
'(* Endof Generate Random Circle *)

'Generate Random Value
Public Function GenerateRandomValue(ByVal Range As Double, Optional Resolution As Double = 10000000#) As Double
  '(* { Result e R As 0 <= Result < Range } *)
  'Randomize
  If Range > 1# Then
    GenerateRandomValue = (1# * RandomL(CLng(Range) - 1)) + ((1# * RandomL(CLng(Resolution))) / Resolution)
    'GenerateRandomValue = Rnd * Range 'RandomS(Range)
  Else
    GenerateRandomValue = ((1# * RandomL(CLng(Resolution))) / Resolution)
  End If
End Function
'(* Endof Generate Random Value *)

'Add
Public Function Add2D(Vec1 As TVector2D, Vec2 As TVector2D) As TVector2D
  Add2D.X = Vec1.X + Vec2.X
  Add2D.Y = Vec1.Y + Vec2.Y
End Function
Public Function Add3D(Vec1 As TVector3D, Vec2 As TVector3D) As TVector3D
  Add3D.X = Vec1.X + Vec2.X
  Add3D.Y = Vec1.Y + Vec2.Y
  Add3D.z = Vec1.z + Vec2.z
End Function
Public Function Add2DA(Vec1 As TVector2DArray, Vec2 As TVector2DArray) As TVector2DArray
Dim i As Long, n As Long, n1 As Long, n2 As Long
  n1 = UBound(Vec1.Arr): n2 = UBound(Vec2.Arr): n = MinL(n1, n2)
  ReDim Add2DA.Arr(n)
  For i = 0 To n
    Add2DA.Arr(i) = Add2D(Vec1.Arr(i), Vec2.Arr(i))
  Next
End Function
Public Function Add3DA(Vec1 As TVector3DArray, Vec2 As TVector3DArray) As TVector3DArray
Dim i As Long, n As Long, n1 As Long, n2 As Long
  n1 = UBound(Vec1.Arr): n2 = UBound(Vec2.Arr): n = MinL(n1, n2)
  ReDim Add3DA.Arr(n)
  For i = 0 To n
    Add3DA.Arr(i) = Add3D(Vec1.Arr(i), Vec2.Arr(i))
  Next
End Function
'(* End of Add *)

'Sub
Public Function Sub2D(Vec1 As TVector2D, Vec2 As TVector2D) As TVector2D
  Sub2D.X = Vec1.X - Vec2.X
  Sub2D.Y = Vec1.Y - Vec2.Y
End Function
Public Function Sub3D(Vec1 As TVector3D, Vec2 As TVector3D) As TVector3D
  Sub3D.X = Vec1.X - Vec2.X
  Sub3D.Y = Vec1.Y - Vec2.Y
  Sub3D.z = Vec1.z - Vec2.z
End Function
Public Function Sub2DA(Vec1 As TVector2DArray, Vec2 As TVector2DArray) As TVector2DArray
Dim i As Long, n As Long, n1 As Long, n2 As Long
  n1 = UBound(Vec1.Arr): n2 = UBound(Vec2.Arr): n = MinL(n1, n2)
  ReDim Sub2DA.Arr(n)
  For i = 0 To n
    Sub2DA.Arr(i) = Sub2D(Vec1.Arr(i), Vec2.Arr(i))
  Next
End Function
Public Function Sub3DA(Vec1 As TVector3DArray, Vec2 As TVector3DArray) As TVector3DArray
Dim i As Long, n As Long, n1 As Long, n2 As Long
  n1 = UBound(Vec1.Arr): n2 = UBound(Vec2.Arr): n = MinL(n1, n2)
  ReDim Sub3DA.Arr(n)
  For i = 0 To n
    Sub3DA.Arr(i) = Sub3D(Vec1.Arr(i), Vec2.Arr(i))
  Next
End Function
'(* End of Sub *)

'Mul
Public Function Mul2D(Vec1 As TVector2D, Vec2 As TVector2D) As TVector3D
  Mul2D.X = 0 'Vec1.y * Vec2.z - Vec1.z * Vec2.y
  Mul2D.Y = 0 'Vec1.z * Vec2.x - Vec1.x * Vec2.z
  Mul2D.z = Vec1.X * Vec2.Y - Vec1.Y * Vec2.X
End Function
Public Function Mul3D(Vec1 As TVector3D, Vec2 As TVector3D) As TVector3D
  Mul3D.X = Vec1.Y * Vec2.z - Vec1.z * Vec2.Y
  Mul3D.Y = Vec1.z * Vec2.X - Vec1.X * Vec2.z
  Mul3D.z = Vec1.X * Vec2.Y - Vec1.Y * Vec2.X
End Function
Public Function Mul3DA(Vec1 As TVector3DArray, Vec2 As TVector3DArray) As TVector3DArray
Dim i As Long, n As Long, n1 As Long, n2 As Long
  n1 = UBound(Vec1.Arr): n2 = UBound(Vec2.Arr): n = MinL(n1, n2)
  ReDim Mul3DA.Arr(n)
  For i = 0 To n
    Mul3DA.Arr(i) = Mul3D(Vec1.Arr(i), Vec2.Arr(i))
  Next
End Function
'(* End of Multiply (cross-product) *)

'UnitVector
Public Function UnitVector2D(Vec As TVector2D) As TVector2D
Dim Mag As Double
  Mag = Magnitude2D(Vec)
  If Mag > 0# Then
    UnitVector2D.X = Vec.X / Mag
    UnitVector2D.Y = Vec.Y / Mag
  Else
    UnitVector2D.X = 0#
    UnitVector2D.Y = 0#
  End If
End Function
Public Function UnitVector3D(Vec As TVector3D) As TVector3D
Dim Mag As Double
  Mag = Magnitude3D(Vec)
  If Mag > 0# Then
    UnitVector3D.X = Vec.X / Mag
    UnitVector3D.Y = Vec.Y / Mag
    UnitVector3D.z = Vec.z / Mag
  Else
    UnitVector3D.X = 0#
    UnitVector3D.Y = 0#
    UnitVector3D.z = 0#
  End If
End Function
'(* End of UnitVector *)

'Magnitude
Public Function Magnitude2D(Vec As TVector2D) As Double
  Magnitude2D = Sqr((Vec.X * Vec.X) + (Vec.Y * Vec.Y))
End Function
Public Function Magnitude3D(Vec As TVector3D) As Double
  Magnitude3D = Sqr((Vec.X * Vec.X) + (Vec.Y * Vec.Y) + (Vec.z * Vec.z))
End Function
'(* End of Magnitude *)

'DotProduct
Public Function DotProduct2D(Vec1 As TVector2D, Vec2 As TVector2D) As Double
  DotProduct2D = Vec1.X * Vec2.X + Vec1.Y * Vec2.Y
End Function
Public Function DotProduct3D(Vec1 As TVector3D, Vec2 As TVector3D) As Double
  DotProduct3D = Vec1.X * Vec2.X + Vec1.Y * Vec2.Y + Vec1.z * Vec2.z
End Function
'(* End of dotProduct *)

'Scale
Public Function Scale2D(Vec As TVector2D, Factor As Double) As TVector2D
  Scale2D.X = Vec.X * Factor
  Scale2D.Y = Vec.Y * Factor
End Function
Public Function Scale3D(Vec As TVector3D, Factor As Double) As TVector3D
  Scale3D.X = Vec.X * Factor
  Scale3D.Y = Vec.Y * Factor
  Scale3D.z = Vec.z * Factor
End Function
Public Function Scale2DA(Vec As TVector2DArray, Factor As Double) As TVector2DArray
Dim i As Long, n As Long
  n = UBound(Vec.Arr)
  ReDim Scale2DA.Arr(n)
  For i = 0 To n
    Scale2DA.Arr(i).X = Vec.Arr(i).X * Factor
    Scale2DA.Arr(i).Y = Vec.Arr(i).Y * Factor
  Next
End Function
Public Function Scale3DA(Vec As TVector3DArray, Factor As Double) As TVector3DArray
Dim i As Long, n As Long
  n = UBound(Vec.Arr)
  ReDim Scale3DA.Arr(n)
  For i = 0 To n
    Scale3DA.Arr(i).X = Vec.Arr(i).X * Factor
    Scale3DA.Arr(i).Y = Vec.Arr(i).Y * Factor
    Scale3DA.Arr(i).z = Vec.Arr(i).z * Factor
  Next
End Function
'(* End of Scale *)

'Negate
Public Function Negate2D(Vec As TVector2D) As TVector2D
  Negate2D.X = -Vec.X
  Negate2D.Y = -Vec.Y
End Function
Public Function Negate3D(Vec As TVector3D) As TVector3D
  Negate3D.X = -Vec.X
  Negate3D.Y = -Vec.Y
  Negate3D.z = -Vec.z
End Function
Public Function Negate2DA(Vec As TVector2DArray) As TVector2DArray
Dim i As Long, n As Long: n = UBound(Vec.Arr)
  ReDim Negate2DA.Arr(n)
  For i = 0 To n
    Negate2DA.Arr(i).X = -Vec.Arr(i).X
    Negate2DA.Arr(i).Y = -Vec.Arr(i).Y
  Next
End Function
Public Function Negate3DA(Vec As TVector3DArray) As TVector3DArray
Dim i As Long, n As Long: n = UBound(Vec.Arr)
  ReDim Negate3DA.Arr(n)
  For i = 0 To n
    Negate3DA.Arr(i).X = -Vec.Arr(i).X
    Negate3DA.Arr(i).Y = -Vec.Arr(i).Y
    Negate3DA.Arr(i).z = -Vec.Arr(i).z
  Next
End Function
'(* End of Negate *)

'IsEqual
Public Function IsEqual(Val1 As Double, Val2 As Double) As Boolean
  IsEqual = IsEqualEps(Val1, Val2, Epsilon)
End Function
Public Function IsEqualEps(Val1 As Double, Val2 As Double, Epsilon As Double) As Boolean
Dim Diff As Double
  Diff = Val1 - Val2
  Call Assert(((-Epsilon <= Diff) And (Diff <= Epsilon)) = (Abs(Diff) <= Epsilon), "Error - Illogical error in equality check. (IsEqual)")
  IsEqualEps = ((-Epsilon <= Diff) And (Diff <= Epsilon))
End Function
Public Function IsEqual2DP(Point1 As TPoint2D, Point2 As TPoint2D) As Boolean
  IsEqual2DP = (IsEqualEps(Point1.X, Point2.X, Epsilon) And IsEqualEps(Point1.Y, Point2.Y, Epsilon))
End Function
Public Function IsEqual2DPEps(Point1 As TPoint2D, Point2 As TPoint2D, Epsilon As Double) As Boolean
  IsEqual2DPEps = (IsEqualEps(Point1.X, Point2.X, Epsilon) And IsEqualEps(Point1.Y, Point2.Y, Epsilon))
End Function
Public Function IsEqual3DP(Point1 As TPoint3D, Point2 As TPoint3D) As Boolean
  IsEqual3DP = (IsEqualEps(Point1.X, Point2.X, Epsilon) And IsEqualEps(Point1.Y, Point2.Y, Epsilon) And IsEqualEps(Point1.z, Point2.z, Epsilon))
End Function
Public Function IsEqual3DPEps(Point1 As TPoint3D, Point2 As TPoint3D, Epsilon As Double) As Boolean
  IsEqual3DPEps = (IsEqualEps(Point1.X, Point2.X, Epsilon) And IsEqualEps(Point1.Y, Point2.Y, Epsilon) And IsEqualEps(Point1.z, Point2.z, Epsilon))
End Function
'(* Endof Is Equal *)

'NotEqual:
Public Function NotEqual(Val1 As Double, Val2 As Double) As Boolean
  NotEqual = NotEqualEps(Val1, Val2, Epsilon)
End Function
Public Function NotEqualEps(Val1 As Double, Val2 As Double, Epsilon As Double) As Boolean
Dim Diff As Double
  Diff = Val1 - Val2
  Call Assert(((-Epsilon > Diff) Or (Diff > Epsilon)) = (Abs(Val1 - Val2) > Epsilon), "Error - Illogical error in equality check. (NotEqual)")
  NotEqualEps = ((-Epsilon > Diff) Or (Diff > Epsilon))
End Function
Public Function NotEqual2DP(Point1 As TPoint2D, Point2 As TPoint2D) As Boolean
  NotEqual2DP = (NotEqualEps(Point1.X, Point2.X, Epsilon) Or NotEqualEps(Point1.Y, Point2.Y, Epsilon))
End Function
Public Function NotEqual2DPEps(Point1 As TPoint2D, Point2 As TPoint2D, Epsilon As Double) As Boolean
  NotEqual2DPEps = (NotEqualEps(Point1.X, Point2.X, Epsilon) Or NotEqualEps(Point1.Y, Point2.Y, Epsilon))
End Function
Public Function NotEqual3DP(Point1 As TPoint3D, Point2 As TPoint3D) As Boolean
  NotEqual3DP = (NotEqualEps(Point1.X, Point2.X, Epsilon) Or NotEqualEps(Point1.Y, Point2.Y, Epsilon) Or NotEqualEps(Point1.z, Point2.z, Epsilon))
End Function
Public Function NotEqual3DPEps(Point1 As TPoint3D, Point2 As TPoint3D, Epsilon As Double) As Boolean
  NotEqual3DPEps = (NotEqualEps(Point1.X, Point2.X, Epsilon) Or NotEqualEps(Point1.Y, Point2.Y, Epsilon) Or NotEqualEps(Point1.z, Point2.z, Epsilon))
End Function
'(* Endof not Equal *)

'LessThanOrEqual
Public Function LessThanOrEqualEps(Val1 As Double, Val2 As Double, Epsilon As Double) As Boolean
  LessThanOrEqualEps = (Val1 < Val2) Or IsEqualEps(Val1, Val2, Epsilon)
End Function
Public Function LessThanOrEqual(Val1 As Double, Val2 As Double) As Boolean
  LessThanOrEqual = (Val1 < Val2) Or IsEqual(Val1, Val2)
End Function
'(* Endof Less Than Or Equal *)

'GreaterThanOrEqual
Public Function GreaterThanOrEqualEps(Val1 As Double, Val2 As Double, Epsilon As Double) As Boolean
  GreaterThanOrEqualEps = (Val1 > Val2) Or IsEqualEps(Val1, Val2, Epsilon)
End Function
Public Function GreaterThanOrEqual(Val1 As Double, Val2 As Double) As Boolean
  GreaterThanOrEqual = (Val1 > Val2) Or IsEqual(Val1, Val2)
End Function
'(* Endof Greater Than Or Equal *)

'IsEqualZero
Public Function IsEqualZeroEps(Val As Double, Epsilon As Double) As Boolean
  IsEqualZeroEps = (Val <= Epsilon)
End Function
Public Function IsEqualZero(Val As Double) As Boolean
  IsEqualZero = IsEqualZeroEps(Val, Epsilon)
End Function
'(* Endof IsEqualZero *)

'IsDegenerate
Public Function IsDegenerate2DXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Boolean
  IsDegenerate2DXY = IsEqual(X1, X2) And IsEqual(Y1, Y2)
End Function
Public Function IsDegenerate2DS(Segment As TSegment2D) As Boolean
  IsDegenerate2DS = IsDegenerate2DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(1).X, Segment.P(1).Y)
End Function
Public Function IsDegenerate2DL(Line As TLine2D) As Boolean
  IsDegenerate2DL = IsDegenerate2DXY(Line.P(0).X, Line.P(0).Y, Line.P(1).X, Line.P(1).Y)
End Function
Public Function IsDegenerate3DXY(X1 As Double, Y1 As Double, z1 As Double, X2 As Double, Y2 As Double, z2 As Double) As Boolean
  IsDegenerate3DXY = IsEqual(X1, X2) And IsEqual(Y1, Y2) And IsEqual(z1, z2)
End Function
Public Function IsDegenerate3DS(Segment As TSegment3D) As Boolean
  IsDegenerate3DS = IsDegenerate3DXY(Segment.P(0).X, Segment.P(0).Y, Segment.P(0).z, Segment.P(1).X, Segment.P(1).Y, Segment.P(1).z)
End Function
Public Function IsDegenerate3DL(Line As TLine3D) As Boolean
  IsDegenerate3DL = IsDegenerate3DXY(Line.P(0).X, Line.P(0).Y, Line.P(0).z, Line.P(1).X, Line.P(1).Y, Line.P(1).z)
End Function
Public Function IsDegenerate2DT(Triangle As TTriangle2D) As Boolean
  IsDegenerate2DT = Collinear2DP(Triangle.P(0), Triangle.P(1), Triangle.P(2)) Or _
            (IsEqual2DP(Triangle.P(0), Triangle.P(1))) Or _
            (IsEqual2DP(Triangle.P(0), Triangle.P(2))) Or _
            (IsEqual2DP(Triangle.P(1), Triangle.P(2)))
End Function
Public Function IsDegenerate3DT(Triangle As TTriangle3D) As Boolean
  IsDegenerate3DT = Collinear3DP(Triangle.P(0), Triangle.P(1), Triangle.P(2)) Or _
            (IsEqual3DP(Triangle.P(0), Triangle.P(1))) Or _
            (IsEqual3DP(Triangle.P(0), Triangle.P(2))) Or _
            (IsEqual3DP(Triangle.P(1), Triangle.P(2)))
End Function
Public Function IsDegenerate2DQ(Quadix As TQuadix2D) As Boolean
           '(* Stage 1 unique points check *)
           '(* Stage 2 collinearity check  *)
  IsDegenerate2DQ = _
            IsDegenerate2DXY(Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(1).X, Quadix.P(1).Y) Or _
            IsDegenerate2DXY(Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(2).X, Quadix.P(2).Y) Or _
            IsDegenerate2DXY(Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(3).X, Quadix.P(3).Y) Or _
            IsDegenerate2DXY(Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(2).X, Quadix.P(2).Y) Or _
            IsDegenerate2DXY(Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(3).X, Quadix.P(3).Y) Or _
            IsDegenerate2DXY(Quadix.P(2).X, Quadix.P(2).Y, Quadix.P(3).X, Quadix.P(3).Y) Or _
            Collinear2DP(Quadix.P(0), Quadix.P(1), Quadix.P(2)) Or _
            Collinear2DP(Quadix.P(1), Quadix.P(2), Quadix.P(3)) Or _
            Collinear2DP(Quadix.P(2), Quadix.P(3), Quadix.P(0)) Or _
            Collinear2DP(Quadix.P(3), Quadix.P(0), Quadix.P(1)) Or _
            Intersect2DP(Quadix.P(0), Quadix.P(1), Quadix.P(2), Quadix.P(3)) Or _
            Intersect2DP(Quadix.P(0), Quadix.P(3), Quadix.P(1), Quadix.P(2)) Or _
            (Not ConvexQuadix(Quadix))
End Function
Public Function IsDegenerate3DQ(Quadix As TQuadix3D) As Boolean
          '(* Stage 1 unique points check *)
          '(* Stage 2 collinearity check  *)
  IsDegenerate3DQ = _
           IsDegenerate3DXY(Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(0).z, Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(1).z) Or _
           IsDegenerate3DXY(Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(0).z, Quadix.P(2).X, Quadix.P(2).Y, Quadix.P(2).z) Or _
           IsDegenerate3DXY(Quadix.P(0).X, Quadix.P(0).Y, Quadix.P(0).z, Quadix.P(3).X, Quadix.P(3).Y, Quadix.P(3).z) Or _
           IsDegenerate3DXY(Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(1).z, Quadix.P(2).X, Quadix.P(2).Y, Quadix.P(2).z) Or _
           IsDegenerate3DXY(Quadix.P(1).X, Quadix.P(1).Y, Quadix.P(1).z, Quadix.P(3).X, Quadix.P(3).Y, Quadix.P(3).z) Or _
           IsDegenerate3DXY(Quadix.P(2).X, Quadix.P(2).Y, Quadix.P(2).z, Quadix.P(3).X, Quadix.P(3).Y, Quadix.P(3).z) Or _
           Collinear3DP(Quadix.P(0), Quadix.P(1), Quadix.P(2)) Or _
           Collinear3DP(Quadix.P(1), Quadix.P(2), Quadix.P(3)) Or _
           Collinear3DP(Quadix.P(2), Quadix.P(3), Quadix.P(0)) Or _
           Collinear3DP(Quadix.P(3), Quadix.P(0), Quadix.P(1))
End Function
Public Function IsDegenerate2DR(Rect As TRectangle) As Boolean
  IsDegenerate2DR = IsEqual2DP(Rect.P(0), Rect.P(1))
End Function
Public Function IsDegenerate2DCi(aCircle As TCircle) As Boolean
  IsDegenerate2DCi = LessThanOrEqual(aCircle.Radius, 0#)
End Function
Public Function IsDegenerate3DSph(Sphere As TSphere) As Boolean
  IsDegenerate3DSph = LessThanOrEqual(Sphere.Radius, 0#)
End Function
Public Function IsDegenerate2DA(Arc As TCircularArc2D) As Boolean
  With Arc
    IsDegenerate2DA = IsDegenerate2DXY(.X1, .Y1, .X2, .Y2) Or _
                      IsDegenerate2DXY(.X1, .Y1, .cX, .cY) Or _
                      IsDegenerate2DXY(.X2, .Y2, .cX, .cY) Or _
                      (LayDistance2DXY(.X1, .Y1, .cX, .cY) <> LayDistance2DXY(.X2, .Y2, .cX, .cY)) Or _
                      (LayDistance2DXY(.X1, .Y1, .cX, .cY) <> LayDistance2DXY(.Px, .Py, .cX, .cY)) Or _
                      (CartesianAngle(.X1 - .cX, .Y1 - .cY) <> .angle1) Or _
                      (CartesianAngle(.X2 - .cX, .Y2 - .cY) <> .angle2) Or _
                      (CartesianAngle(.Px - .cX, .Py - .cY) <> Abs(.angle1 - .angle2)) Or _
                      (Orientation2D(.X1, .Y1, .X2, .Y2, .Px, .Py) <> .Orientation)
   End With
End Function
Public Function IsDegenerateO(Obj As TGeometricObject) As Boolean
  Select Case Obj.ObjectType
  Case geoSegment2D:  IsDegenerateO = IsDegenerate2DS(Obj.Segment2D)
  Case geoSegment3D:  IsDegenerateO = IsDegenerate3DS(Obj.Segment3D)
  Case geoLine2D:     IsDegenerateO = IsDegenerate2DL(Obj.Line2D)
  Case geoLine3D:     IsDegenerateO = IsDegenerate3DL(Obj.Line3D)
  Case geoTriangle2D: IsDegenerateO = IsDegenerate2DT(Obj.Triangle2D)
  Case geoTriangle3D: IsDegenerateO = IsDegenerate3DT(Obj.Triangle3D)
  Case geoQuadix2D:   IsDegenerateO = IsDegenerate2DQ(Obj.Quadix2D)
  Case geoQuadix3D:   IsDegenerateO = IsDegenerate3DQ(Obj.Quadix3D)
  Case geoRectangle:  IsDegenerateO = IsDegenerate2DR(Obj.Rectangle)
  Case geoCircle:     IsDegenerateO = IsDegenerate2DCi(Obj.aCircle)
  Case geoSphere:     IsDegenerateO = IsDegenerate3DSph(Obj.Sphere)
  Case Else:          IsDegenerateO = False
  End Select
End Function
'(* Endof IsDegenerate *)

'Swap
Public Sub SwapS(Val1 As Double, Val2 As Double)
Dim Temp As Double
  Temp = Val1
  Val1 = Val2
  Val2 = Temp
End Sub
Public Sub SwapL(Val1 As Long, Val2 As Long)
Dim Temp As Long
  Temp = Val1
  Val1 = Val2
  Val2 = Temp
End Sub
Public Sub Swap2DP(Point1 As TPoint2D, Point2 As TPoint2D)
  Call SwapS(Point1.X, Point2.X)
  Call SwapS(Point1.Y, Point2.Y)
End Sub
Public Sub Swap3DP(Point1 As TPoint3D, Point2 As TPoint3D)
  Call SwapS(Point1.X, Point2.X)
  Call SwapS(Point1.Y, Point2.Y)
  Call SwapS(Point1.z, Point2.z)
End Sub
Public Sub Swap2DS(Segment1 As TSegment2D, Segment2 As TSegment2D)
  Call Swap2DP(Segment1.P(0), Segment2.P(0))
  Call Swap2DP(Segment1.P(1), Segment2.P(1))
End Sub
Public Sub Swap3DS(Segment1 As TSegment3D, Segment2 As TSegment3D)
  Call Swap3DP(Segment1.P(0), Segment2.P(0))
  Call Swap3DP(Segment1.P(1), Segment2.P(1))
End Sub
Public Sub Swap2DL(Line1 As TLine2D, Line2 As TLine2D)
  Call Swap2DP(Line1.P(0), Line2.P(0))
  Call Swap2DP(Line1.P(1), Line2.P(1))
End Sub
Public Sub Swap2DT(Triangle1 As TTriangle2D, Triangle2 As TTriangle2D)
  Call Swap2DP(Triangle1.P(0), Triangle2.P(0))
  Call Swap2DP(Triangle1.P(1), Triangle2.P(1))
  Call Swap2DP(Triangle1.P(2), Triangle2.P(2))
End Sub
Public Sub Swap3DT(Triangle1 As TTriangle3D, Triangle2 As TTriangle3D)
  Call Swap3DP(Triangle1.P(0), Triangle2.P(0))
  Call Swap3DP(Triangle1.P(1), Triangle2.P(1))
  Call Swap3DP(Triangle1.P(2), Triangle2.P(2))
End Sub
Public Sub Swap2DQ(Quadix1 As TQuadix2D, Quadix2 As TQuadix2D)
  Call Swap2DP(Quadix1.P(0), Quadix2.P(0))
  Call Swap2DP(Quadix1.P(1), Quadix2.P(1))
  Call Swap2DP(Quadix1.P(2), Quadix2.P(2))
  Call Swap2DP(Quadix1.P(3), Quadix2.P(3))
End Sub
Public Sub Swap3DQ(Quadix1 As TQuadix3D, Quadix2 As TQuadix3D)
  Call Swap3DP(Quadix1.P(0), Quadix2.P(0))
  Call Swap3DP(Quadix1.P(1), Quadix2.P(1))
  Call Swap3DP(Quadix1.P(2), Quadix2.P(2))
  Call Swap3DP(Quadix1.P(3), Quadix2.P(3))
End Sub
Public Sub SwapC(Circle1 As TCircle, Circle2 As TCircle)
  Call SwapS(Circle1.X, Circle2.X)
  Call SwapS(Circle1.Y, Circle2.Y)
  Call SwapS(Circle1.Radius, Circle2.Radius)
End Sub
Public Sub SwapSph(Sphere1 As TSphere, Sphere2 As TSphere)
  Call SwapS(Sphere1.X, Sphere2.X)
  Call SwapS(Sphere1.Y, Sphere2.Y)
  Call SwapS(Sphere1.z, Sphere2.z)
  Call SwapS(Sphere1.Radius, Sphere2.Radius)
End Sub
Public Sub Swap2DA(Arc1 As TCircularArc2D, Arc2 As TCircularArc2D)
  Call SwapS(Arc1.X1, Arc2.X1)
  Call SwapS(Arc1.X2, Arc2.X2)
  Call SwapS(Arc1.cX, Arc2.cX)
  Call SwapS(Arc1.Px, Arc2.Px)
  Call SwapS(Arc1.Y1, Arc2.Y1)
  Call SwapS(Arc1.Y2, Arc2.Y2)
  Call SwapS(Arc1.cY, Arc2.cY)
  Call SwapS(Arc1.Py, Arc2.Py)
  Call SwapS(Arc1.angle1, Arc2.angle1)
  Call SwapS(Arc1.angle2, Arc2.angle2)
  Call SwapL(Arc1.Orientation, Arc2.Orientation)
End Sub
'(* Endof Swap *)

'ZeroEquivalency
Public Function ZeroEquivalency() As Boolean
  ZeroEquivalency = IsEqual(CalculateSystemEpsilon, 0#)
End Function
'(* Endof ZeroEquivalency *)

Public Function ExecuteTests() As TNumericPrecisionResult
  With ExecuteTests
    .EEResult = LessThanOrEqual(SystemEpsilon, Epsilon)
    .ZEResult = ZeroEquivalency
    .EFPResult = ExtendedFloatingPointTest
    .SystemEpsilon = SystemEpsilon
  End With
End Function
'(* Endof ExecuteTests *)

'RotationTest
Public Function RotationTest(RadiusLength As Double) As Boolean
Const delta_angle = 0.036
Dim InitialPoint As TPoint2D, Point As TPoint2D, TempPoint As TPoint2D, i As Long
  InitialPoint = EquatePoint2D(RadiusLength, 0)
  Point = InitialPoint
  For i = 1 To 10000
    TempPoint = Rotate2DP(delta_angle, Point)
    Point = TempPoint
  Next
  RotationTest = IsEqual2DP(InitialPoint, Point)
End Function
'(* Endof RotationTest *)

'ExtendedFloatingPointTest
Public Function ExtendedFloatingPointTest() As Boolean
  '(*
  '  Yet to be completed! **************************
  '*)
Dim Large_Radius As Double, Small_Radius As Double
  Large_Radius = 100000#
  Small_Radius = 100000# 'warum genauso groß ???????
  ExtendedFloatingPointTest = RotationTest(Large_Radius) And RotationTest(Small_Radius)
End Function
'(* Endof ExtendedFloatingPointTest *)

'CalculateSystemEpsilon
Private Function CalculateSystemEpsilon() As Double
Dim Epsilon As Double
Dim Check As Double
Dim LastCheck As Double
  Epsilon = 1#
  Check = 1#
  Do
    LastCheck = Check
    Epsilon = Epsilon * 0.5
    Check = 1# + Epsilon
  Loop Until (Check = 1#) Or (Check = LastCheck)
  CalculateSystemEpsilon = Epsilon
End Function
'(* Endof CalculateSystemEpsilon *)

Public Function Assert(BolVal As Boolean, aMess As String)
  If Not BolVal Then MsgBox aMess
End Function

'RandomValue
Public Function RandomValue(Optional ResInt As Long = RandomResolutionInt, Optional ResFlt As Double = RandomResolutionFlt) As Double
  RandomValue = (1# * (Rnd * ResInt)) / ResFlt
End Function
'(* Endof Random Value *)

'InitialiseTrigonometryTables
Private Sub InitialiseTrigonometryTables()
Dim i As Long
' (*
'    Note: Trigonometry look-up tables are used to speed-up
'          sine, cosine and tangent calculations.
' *)
  ReDim CosTable(360 - 1)
  ReDim SinTable(360 - 1)
  ReDim TanTable(360 - 1)
  For i = 0 To 359
    CosTable(i) = Cos((1# * i) * PIDiv180)
    SinTable(i) = Sin((1# * i) * PIDiv180)
    TanTable(i) = Tan((1# * i) * PIDiv180)
  Next
End Sub
'(* End of Initialise Trigonometry Tables *)
'hier noch zusätzlich ConvexHull und OrderedPolygon
Public Function ConvexHull2D(Points() As TPoint2D) As TPolygon2D
Dim aConvexHull2D As New TConvexHull2D
#If defFastGeoAXdll Then
  Set aConvexHull2D.FastGeo = Me
#End If
  ConvexHull2D = aConvexHull2D.ConvexHull(Points)
End Function
'Public Function ConvexHull3D(Points() As TPoint3D) As TPolygon3D
'Dim aConvexHull3D As New TConvexHull3D
'  ConvexHull = aConvexHull2D.ConvexHull(Points)
'End Function
Public Sub OrderedPolygon2D(Points() As TPoint2D)
Dim aOrderedPolygon As New TOrderedPolygon2D
#If defFastGeoAXdll Then
  Set aOrderedPolygon.FastGeo = Me
#End If
  Call aOrderedPolygon.OrderedPolygon(Points())
End Sub
Public Sub OrderedPolygon3D(Points() As TPoint3D)
Dim aOrderedPolygon As New TOrderedPolygon3D
#If defFastGeoAXdll Then
  Set aOrderedPolygon.FastGeo = Me
#End If
  Call aOrderedPolygon.OrderedPolygon(Points())
End Sub
'###########################################################################

Public Sub Initialize()
  SystemEpsilon = CalculateSystemEpsilon
  MaximumX = Infinity
  MinimumX = -Infinity
  MaximumY = Infinity
  MinimumY = -Infinity
  MaximumZ = Infinity
  MinimumZ = -Infinity
  InitialiseTrigonometryTables
'  VersionInformation = "FastGEO Version 5.0.1"
'  AuthorInformation = "Arash Partow (1997-2006)"
'  EpochInformation = "Lambda-Phi"
'  RecentUpdate = "10-02-2006"
'  LCID = "$10-02-2006:FEDEDB4632780C2FAE$"
End Sub
'(* Endof Initialise *)

Public Sub Terminate()
  Erase CosTable
  Erase SinTable
  Erase TanTable
End Sub
