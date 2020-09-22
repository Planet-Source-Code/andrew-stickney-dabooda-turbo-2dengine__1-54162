VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBAutoMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetWidth(Value As Single)
    AutoMove.Width = Value
End Sub
Public Sub SetHeight(Value As Single)
    AutoMove.Height = Value
End Sub
Public Sub SetSWidth(Value As Single)
    AutoMove.SWidth = Value / 2
End Sub
Public Sub SetSHeight(Value As Single)
    AutoMove.SHeight = Value / 2
End Sub
Public Sub SetOn(Value As Boolean)
    AutoMove.On = Value
End Sub
Public Sub SetMapReference(Value As Integer)
    AutoMove.Mref = Value
End Sub
Public Sub SetSpriteReference(Value As Integer)
    AutoMove.Sref = Value
End Sub

Public Function GetWidth() As Single
    GetWidth = AutoMove.Width
End Function
Public Function GetHeight() As Single
    GetHeight = AutoMove.Height
End Function
Public Function GetSWidth() As Single
    GetSWidth = AutoMove.SWidth * 2
End Function
Public Function GetSHeight() As Single
    GetSHeight = AutoMove.SHeight * 2
End Function
Public Function GetOn() As Boolean
    GetOn = AutoMove.On
End Function
Public Function GetSpriteReference() As Integer
    GetSpriteReference = AutoMove.Sref
End Function

Private Sub Class_Initialize()
    With AutoMove
        .Width = 64
        .Height = 64
        .SHeight = 16
        .SWidth = 16
        .Left = True
        .Right = True
        .Up = True
        .Down = True
        .On = False
        .Sref = 0
    End With
End Sub
