VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBSubMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetTextureReference(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).TextureRef = Value
End Sub

Public Sub SetULAlpha(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).ULCol.Alpha = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetULRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).ULCol.Red = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetULGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).ULCol.Green = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetULBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).ULCol.Blue = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub

Public Sub SetURAlpha(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).URCol.Alpha = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetURRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).URCol.Red = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetURGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).URCol.Green = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetURBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).URCol.Blue = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub

Public Sub SetLLAlpha(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LLCol.Alpha = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetLLRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LLCol.Red = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetLLGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LLCol.Green = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetLLBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LLCol.Blue = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub

Public Sub SetLRAlpha(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LRCol.Alpha = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetLRRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LRCol.Red = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetLRGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LRCol.Green = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub
Public Sub SetLRBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LRCol.Blue = Value
    UpdateColor MIndex, XIndex, YIndex
End Sub


Public Sub SetULSRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).ULSpec.Red = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub
Public Sub SetULSGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).ULSpec.Green = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub
Public Sub SetULSBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).ULSpec.Blue = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub

Public Sub SetURSRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).URSpec.Red = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub
Public Sub SetURSGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).URSpec.Green = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub
Public Sub SetURSBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).URSpec.Blue = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub

Public Sub SetLLSRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LLSpec.Red = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub
Public Sub SetLLSGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LLSpec.Green = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub
Public Sub SetLLSBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LLSpec.Blue = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub

Public Sub SetLRSRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LRSpec.Red = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub
Public Sub SetLRSGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LRSpec.Green = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub
Public Sub SetLRSBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    SubMap(MIndex, XIndex, YIndex).LRSpec.Blue = Value
    UpdateSpecular MIndex, XIndex, YIndex
End Sub

Public Sub SetVisible(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Boolean)
    SubMap(MIndex, XIndex, YIndex).Visible = Value
End Sub

Public Sub SetLimitUp(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Boolean)
    SubMap(MIndex, XIndex, YIndex).LUp = Value
End Sub
Public Sub SetLimitDown(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Boolean)
    SubMap(MIndex, XIndex, YIndex).LDown = Value
End Sub
Public Sub SetLimitLeft(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Boolean)
    SubMap(MIndex, XIndex, YIndex).LLeft = Value
End Sub
Public Sub SetLimitRight(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Boolean)
    SubMap(MIndex, XIndex, YIndex).LRight = Value
End Sub

Public Function GetTextureReference(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetTextureReference = SubMap(MIndex, XIndex, YIndex).TextureRef
End Function

Public Function GetULAlpha(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetULAlpha = SubMap(MIndex, XIndex, YIndex).ULCol.Alpha
End Function
Public Function GetULRed(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetULRed = SubMap(MIndex, XIndex, YIndex).ULCol.Red
End Function
Public Function GetULGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetULGreen = SubMap(MIndex, XIndex, YIndex).ULCol.Green
End Function
Public Function GetULBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetULBlue = SubMap(MIndex, XIndex, YIndex).ULCol.Blue
End Function

Public Function GetURAlpha(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetURAlpha = SubMap(MIndex, XIndex, YIndex).URCol.Alpha
End Function
Public Function GetURRed(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetURRed = SubMap(MIndex, XIndex, YIndex).URCol.Red
End Function
Public Function GetURGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetURGreen = SubMap(MIndex, XIndex, YIndex).URCol.Green
End Function
Public Function GetURBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetURBlue = SubMap(MIndex, XIndex, YIndex).URCol.Blue
End Function

Public Function GetLLAlpha(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLLAlpha = SubMap(MIndex, XIndex, YIndex).LLCol.Alpha
End Function
Public Function GetLLRed(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLLRed = SubMap(MIndex, XIndex, YIndex).LLCol.Red
End Function
Public Function GetLLGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLLGreen = SubMap(MIndex, XIndex, YIndex).LLCol.Green
End Function
Public Function GetLLBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLLBlue = SubMap(MIndex, XIndex, YIndex).LLCol.Blue
End Function

Public Function GetLRAlpha(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLRAlpha = SubMap(MIndex, XIndex, YIndex).LRCol.Alpha
End Function
Public Function GetLRRed(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLRRed = SubMap(MIndex, XIndex, YIndex).LRCol.Red
End Function
Public Function GetLRGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLRGreen = SubMap(MIndex, XIndex, YIndex).LRCol.Green
End Function
Public Function GetLRBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLRBlue = SubMap(MIndex, XIndex, YIndex).LRCol.Blue
End Function

Public Function GetULSRed(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetULSRed = SubMap(MIndex, XIndex, YIndex).ULSpec.Red
End Function
Public Function GetULSGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetULSGreen = SubMap(MIndex, XIndex, YIndex).ULSpec.Green
End Function
Public Function GetULSBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetULSBlue = SubMap(MIndex, XIndex, YIndex).ULSpec.Blue
End Function

Public Function GetURSRed(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetURSRed = SubMap(MIndex, XIndex, YIndex).URSpec.Red
End Function
Public Function GetURSGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetURSGreen = SubMap(MIndex, XIndex, YIndex).URSpec.Green
End Function
Public Function GetURSBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetURSBlue = SubMap(MIndex, XIndex, YIndex).URSpec.Blue
End Function

Public Function GetLLSRed(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLLSRed = SubMap(MIndex, XIndex, YIndex).LLSpec.Red
End Function
Public Function GetLLSGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLLSGreen = SubMap(MIndex, XIndex, YIndex).LLSpec.Green
End Function
Public Function GetLLSBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLLSBlue = SubMap(MIndex, XIndex, YIndex).LLSpec.Blue
End Function

Public Function GetLRSRed(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLRSRed = SubMap(MIndex, XIndex, YIndex).LRSpec.Red
End Function
Public Function GetLRSGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLRSGreen = SubMap(MIndex, XIndex, YIndex).LRSpec.Green
End Function
Public Function GetLRSBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Integer
    GetLRSBlue = SubMap(MIndex, XIndex, YIndex).LRSpec.Blue
End Function

Public Function GetVisible(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Boolean
    GetVisible = SubMap(MIndex, XIndex, YIndex).Visible
End Function

Public Function GetLimitUp(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Boolean
    GetLimitUp = SubMap(MIndex, XIndex, YIndex).LUp
End Function
Public Function GetLimitDown(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Boolean
    GetLimitDown = SubMap(MIndex, XIndex, YIndex).LDown
End Function
Public Function GetLimitLeft(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Boolean
    GetLimitLeft = SubMap(MIndex, XIndex, YIndex).LLeft
End Function
Public Function GetLimitRight(MIndex As Integer, XIndex As Integer, YIndex As Integer) As Boolean
    GetLimitRight = SubMap(MIndex, XIndex, YIndex).LRight
End Function

Public Sub QMSetAlpha(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        .ULCol.Alpha = Value
        .URCol.Alpha = Value
        .LLCol.Alpha = Value
        .LRCol.Alpha = Value
    End With
    UpdateColor MIndex, XIndex, YIndex
End Sub

Public Sub QMSetRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        .ULCol.Red = Value
        .URCol.Red = Value
        .LLCol.Red = Value
        .LRCol.Red = Value
    End With
    UpdateColor MIndex, XIndex, YIndex
End Sub

Public Sub QMSetGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        .ULCol.Green = Value
        .URCol.Green = Value
        .LLCol.Green = Value
        .LRCol.Green = Value
    End With
    UpdateColor MIndex, XIndex, YIndex
End Sub

Public Sub QMSetBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        .ULCol.Blue = Value
        .URCol.Blue = Value
        .LLCol.Blue = Value
        .LRCol.Blue = Value
    End With
    UpdateColor MIndex, XIndex, YIndex
End Sub

Public Sub QMSetSRed(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        .ULSpec.Red = Value
        .URSpec.Red = Value
        .LLSpec.Red = Value
        .LRSpec.Red = Value
    End With
    UpdateSpecular MIndex, XIndex, YIndex
End Sub

Public Sub QMSetSGreen(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        .ULSpec.Green = Value
        .URSpec.Green = Value
        .LLSpec.Green = Value
        .LRSpec.Green = Value
    End With
    UpdateSpecular MIndex, XIndex, YIndex
End Sub

Public Sub QMSetSBlue(MIndex As Integer, XIndex As Integer, YIndex As Integer, Value As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        .ULSpec.Blue = Value
        .URSpec.Blue = Value
        .LLSpec.Blue = Value
        .LRSpec.Blue = Value
    End With
    UpdateSpecular MIndex, XIndex, YIndex
End Sub

Public Sub QMSetLimit(MIndex As Integer, XIndex As Integer, YIndex As Integer, Up As Boolean, Left As Boolean, Down As Boolean, Right As Boolean)
    With SubMap(MIndex, XIndex, YIndex)
        .LUp = Up
        .LLeft = Left
        .LDown = Down
        .LRight = Right
    End With
End Sub

Public Sub QMSetColor(MIndex As Integer, XIndex As Integer, YIndex As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        With .ULCol
            .Red = Red
            .Green = Green
            .Blue = Blue
        End With
        With .URCol
            .Red = Red
            .Green = Green
            .Blue = Blue
        End With
        With .LLCol
            .Red = Red
            .Green = Green
            .Blue = Blue
        End With
        With .LRCol
            .Red = Red
            .Green = Green
            .Blue = Blue
        End With
    End With
    UpdateColor MIndex, XIndex, YIndex
End Sub

Public Sub QMSetSpecular(MIndex As Integer, XIndex As Integer, YIndex As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        With .ULSpec
            .Red = Red
            .Green = Green
            .Blue = Blue
        End With
        With .URSpec
            .Red = Red
            .Green = Green
            .Blue = Blue
        End With
        With .LLSpec
            .Red = Red
            .Green = Green
            .Blue = Blue
        End With
        With .LRSpec
            .Red = Red
            .Green = Green
            .Blue = Blue
        End With
    End With
End Sub

Private Sub Class_Initialize()
Dim A As Integer, B As Integer, T As Integer
    For T = 1 To 8
        For A = 1 To 64
        For B = 1 To 64
            With SubMap(T, A, B)
                .TextureRef = 0
                .ULCol.Alpha = 255
                .ULCol.Red = 255
                .ULCol.Green = 255
                .ULCol.Blue = 255
                .URCol.Alpha = 255
                .URCol.Red = 255
                .URCol.Green = 255
                .URCol.Blue = 255
                .LLCol.Alpha = 255
                .LLCol.Red = 255
                .LLCol.Green = 255
                .LLCol.Blue = 255
                .LRCol.Alpha = 255
                .LRCol.Red = 255
                .LRCol.Green = 255
                .LRCol.Blue = 255
                .ULSpec.Alpha = 255
                .ULSpec.Red = 0
                .ULSpec.Green = 0
                .ULSpec.Blue = 0
                .URSpec.Alpha = 255
                .URSpec.Red = 0
                .URSpec.Green = 0
                .URSpec.Blue = 0
                .LLSpec.Alpha = 255
                .LLSpec.Red = 0
                .LLSpec.Green = 0
                .LLSpec.Blue = 0
                .LRSpec.Alpha = 255
                .LRSpec.Red = 0
                .LRSpec.Green = 0
                .LRSpec.Blue = 0
                .Visible = True
                .LUp = False
                .LDown = False
                .LLeft = False
                .LRight = False
                .C1 = D3DColorARGB(255, 255, 255, 255)
                .C2 = D3DColorARGB(255, 255, 255, 255)
                .C3 = D3DColorARGB(255, 255, 255, 255)
                .C4 = D3DColorARGB(255, 255, 255, 255)
                .S1 = D3DColorARGB(255, 0, 0, 0)
                .S2 = D3DColorARGB(255, 0, 0, 0)
                .S3 = D3DColorARGB(255, 0, 0, 0)
                .S4 = D3DColorARGB(255, 0, 0, 0)
            End With
        Next B
        Next A
    Next T
End Sub

Private Sub UpdateColor(MIndex As Integer, XIndex As Integer, YIndex As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        .C1 = D3DColorARGB(.ULCol.Alpha, .ULCol.Red, .ULCol.Green, .ULCol.Blue)
        .C2 = D3DColorARGB(.URCol.Alpha, .URCol.Red, .URCol.Green, .URCol.Blue)
        .C3 = D3DColorARGB(.LLCol.Alpha, .LLCol.Red, .LLCol.Green, .LLCol.Blue)
        .C4 = D3DColorARGB(.LRCol.Alpha, .LRCol.Red, .LRCol.Green, .LRCol.Blue)
    End With
End Sub

Private Sub UpdateSpecular(MIndex As Integer, XIndex As Integer, YIndex As Integer)
    With SubMap(MIndex, XIndex, YIndex)
        .S1 = D3DColorARGB(255, .ULSpec.Red, .ULSpec.Green, .ULSpec.Blue)
        .S2 = D3DColorARGB(255, .URSpec.Red, .URSpec.Green, .URSpec.Blue)
        .S3 = D3DColorARGB(255, .LLSpec.Red, .LLSpec.Green, .LLSpec.Blue)
        .S4 = D3DColorARGB(255, .LRSpec.Red, .LRSpec.Green, .LRSpec.Blue)
    End With
End Sub
