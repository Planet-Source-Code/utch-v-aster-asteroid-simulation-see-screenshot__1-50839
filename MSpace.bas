Attribute VB_Name = "MSpace"
Option Explicit

Public Type Rock
    Radius As Integer
    Speed As Integer
    XSlope As Integer
    YSlope As Integer
    XSpot As Integer
    YSpot As Integer
    XStart As Integer
    YStart As Integer
End Type
Public Rocks(0 To 999) As Rock

Public Function RandomNum(ByVal lMin As Long, ByVal lMax As Long)
    Randomize (Second(Time$))
    RandomNum = Int((Rnd * lMax) + lMin)
End Function

