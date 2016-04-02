Attribute VB_Name = "ModuleElement"
Option Explicit

Public Enum EleStyle
    Vol
    Cur
    Res
End Enum

Public Type Element
    Node1 As Integer
    Node2 As Integer
    Style As EleStyle
    Value As Double
End Type

Public Eles(32767) As Element
Public EleCount As Integer

Public EleStyleName(2) As String
Public EleStyleUnit(2) As String
