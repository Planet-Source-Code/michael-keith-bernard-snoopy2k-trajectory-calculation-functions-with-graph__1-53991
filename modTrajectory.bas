Attribute VB_Name = "modTrajectory"
Option Explicit

'**************************************************
'**************************************************
'Name:        modTrajectory.bas
'Type/Build:  Visual Basic 6.0 (SP4), Module/Class Module
'Author:      Michael-Keith Bernard (Snoopy2K)
'Date:        Sunday, May 23, 2004
'Purpose:     Several Trajectory Calculations.
'Note:        Equation Source: http://hyperphysics.phy-astr.gsu.edu/hbase/traj.html
'**************************************************
'**************************************************

Public Const EARTHS_GRAVITY As Double = 9.80665

Public Function ConvertToRadian(ByVal dblDegree As Double) As Double
Dim PI As Double
PI = Atn(1) * 4
    ConvertToRadian = CDbl(dblDegree * (PI / 180))
End Function

Public Function ConvertToDegree(ByVal dblRadian As Double) As Double
Dim PI As Double
PI = Atn(1) * 4
    ConvertToDegree = CDbl(dblRadian * (180 / PI))
End Function

Public Function VeloX(ByVal dblVelocity As Double, ByVal dblAngle As Double) As Double
    VeloX = CDbl(dblVelocity * Cos(ConvertToRadian(dblAngle)))
End Function

Public Function VeloY(ByVal dblVelocity As Double, ByVal dblAngle As Double) As Double
    VeloY = CDbl(dblVelocity * Sin(ConvertToRadian(dblAngle)))
End Function

Public Function Apex(ByVal dblVelocity As Double, ByVal dblAngle As Double, ByVal dblGravity As Double) As Double
    Apex = CDbl((dblVelocity ^ 2 * (Sin(ConvertToRadian(dblAngle)) ^ 2)) / (2 * dblGravity))
End Function

Public Function Range(ByVal dblVelocity As Double, ByVal dblAngle As Double, ByVal dblGravity As Double) As Double
    Range = CDbl((dblVelocity ^ 2 * (Sin(ConvertToRadian(2 * dblAngle)))) / dblGravity)
End Function

Public Function AirTime(ByVal dblVelocity As Double, ByVal dblAngle As Double, ByVal dblGravity As Double) As Double
    AirTime = CDbl(((2 * dblVelocity) * (Sin(ConvertToRadian(dblAngle)))) / dblGravity)
End Function

Public Function XPosAtTime(ByVal dblVelocity As Double, ByVal dblAngle As Double, ByVal dblTime As Double) As Double
    XPosAtTime = CDbl(VeloX(dblVelocity, dblAngle) * dblTime)
End Function

Public Function YPosAtTime(ByVal dblVelocity As Double, ByVal dblAngle As Double, ByVal dblGravity As Double, ByVal dblTime As Double) As Double
    YPosAtTime = CDbl((VeloY(dblVelocity, dblAngle) * dblTime) - (0.5 * dblGravity * (dblTime ^ 2)))
End Function
