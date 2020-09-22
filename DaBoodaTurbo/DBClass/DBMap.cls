VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetXCount(Index As Integer, Value As Integer)
    Map(Index).XCount = Value
    Map(Index).Width = Map(Index).XCount * Map(Index).SubMapWidth
End Sub
Public Sub SetYCount(Index As Integer, Value As Integer)
    Map(Index).YCount = Value
    Map(Index).Height = Map(Index).YCount * Map(Index).SubMapHeight
End Sub
Public Sub SetXPosition(Index As Integer, Value As Single)
    Map(Index).XPos = Value
End Sub
Public Sub SetYPosition(Index As Integer, Value As Single)
    Map(Index).YPos = Value
End Sub
Public Sub SetXIncrement(Index As Integer, Value As Single)
    Map(Index).XInc = Value
End Sub
Public Sub SetYIncrement(Index As Integer, Value As Single)
    Map(Index).YInc = Value
End Sub
Public Sub SetXReference(Index As Integer, Value As Integer)
    Map(Index).XRef = Value
End Sub
Public Sub SetYReference(Index As Integer, Value As Integer)
    Map(Index).YRef = Value
End Sub
Public Sub SetLooping(Index As Integer, Value As Boolean)
    Map(Index).Looping = Value
End Sub
Public Sub SetVisible(Index As Integer, Value As Boolean)
    Map(Index).Visible = Value
End Sub
Public Sub SetSubMapWidth(Index As Integer, Value As Single)
    Map(Index).SubMapWidth = Value
    Map(Index).Width = Map(Index).XCount * Map(Index).SubMapWidth
End Sub
Public Sub SetSubMapHeight(Index As Integer, Value As Single)
    Map(Index).SubMapHeight = Value
    Map(Index).Height = Map(Index).YCount * Map(Index).SubMapHeight
End Sub
Public Sub SetZOrder(Index As Integer, Value As Integer)
    Map(Index).ZOrder = Value
End Sub

Public Function GetXCount(Index As Integer) As Integer
    GetXCount = Map(Index).XCount
End Function
Public Function GetYCount(Index As Integer) As Integer
    GetYCount = Map(Index).YCount
End Function
Public Function GetXPosition(Index As Integer) As Single
    GetXPosition = Map(Index).XPos
End Function
Public Function GetYPosition(Index As Integer) As Single
    GetYPosition = Map(Index).YPos
End Function
Public Function GetXIncrement(Index As Integer) As Single
    GetXIncrement = Map(Index).XPos
End Function
Public Function GetYIncrement(Index As Integer) As Single
    GetYIncrement = Map(Index).YInc
End Function
Public Function GetXReference(Index As Integer) As Integer
    GetXReference = Map(Index).XRef
End Function
Public Function GetYReference(Index As Integer) As Integer
    GetYReference = Map(Index).YRef
End Function
Public Function GetLooping(Index As Integer) As Boolean
    GetLooping = Map(Index).Looping
End Function
Public Function GetVisible(Index As Integer) As Boolean
    GetVisible = Map(Index).Visible
End Function
Public Function GetSubMapWidth(Index As Integer) As Single
    GetSubMapWidth = Map(Index).SubMapWidth
End Function
Public Function GetSubMapHeight(Index As Integer) As Single
    GetSubMapHeight = Map(Index).SubMapHeight
End Function
Public Function GetZOrder(Index As Integer) As Integer
    GetZOrder = Map(Index).ZOrder
End Function

Public Sub QMSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
Dim A As Integer, B As Integer
    For A = 1 To Map(Index).XCount
    For B = 1 To Map(Index).YCount
        With SubMap(Index, A, B)
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
            .C1 = D3DColorARGB(.ULCol.Alpha, .ULCol.Red, .ULCol.Green, .ULCol.Blue)
            .C2 = D3DColorARGB(.URCol.Alpha, .URCol.Red, .URCol.Green, .URCol.Blue)
            .C3 = D3DColorARGB(.LLCol.Alpha, .LLCol.Red, .LLCol.Green, .LLCol.Blue)
            .C4 = D3DColorARGB(.LRCol.Alpha, .LRCol.Red, .LRCol.Green, .LRCol.Blue)
        End With
    Next B
    Next A
End Sub

Public Sub QMSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
Dim A As Integer, B As Integer
    For A = 1 To Map(Index).XCount
    For B = 1 To Map(Index).YCount
        With SubMap(Index, A, B)
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
            .S1 = D3DColorARGB(255, .ULSpec.Red, .ULSpec.Green, .ULSpec.Blue)
            .S2 = D3DColorARGB(255, .URSpec.Red, .URSpec.Green, .URSpec.Blue)
            .S3 = D3DColorARGB(255, .LLSpec.Red, .LLSpec.Green, .LLSpec.Blue)
            .S4 = D3DColorARGB(255, .LRSpec.Red, .LRSpec.Green, .LRSpec.Blue)
        End With
    Next B
    Next A
End Sub

Public Sub QMSetAlpha(Index As Integer, Value As Integer)
Dim A As Integer, B As Integer
    For A = 1 To Map(Index).XCount
    For B = 1 To Map(Index).YCount
        With SubMap(Index, A, B)
            .ULCol.Alpha = Value
            .URCol.Alpha = Value
            .LLCol.Alpha = Value
            .LRCol.Alpha = Value
            .C1 = D3DColorARGB(.ULCol.Alpha, .ULCol.Red, .ULCol.Green, .ULCol.Blue)
            .C2 = D3DColorARGB(.URCol.Alpha, .URCol.Red, .URCol.Green, .URCol.Blue)
            .C3 = D3DColorARGB(.LLCol.Alpha, .LLCol.Red, .LLCol.Green, .LLCol.Blue)
            .C4 = D3DColorARGB(.LRCol.Alpha, .LRCol.Red, .LRCol.Green, .LRCol.Blue)
        End With
    Next B
    Next A
End Sub

Public Sub QMSetGetRectangle(Index As Integer, Size As Single, X As Single, Y As Single, Width As Single, Height As Single)
    With Map(Index)
        .Tx1 = TextureValue(Size, X)
        .Tx2 = TextureValue(Size, X + Width)
        .Ty1 = TextureValue(Size, Y)
        .Ty2 = TextureValue(Size, Y + Height)
    End With
End Sub

Public Sub QMResetMap(Index As Integer)
Dim A As Integer, B As Integer
        With Map(Index)
            .XCount = 64
            .YCount = 64
            .XPos = 0
            .YPos = 0
            .XInc = 1
            .YInc = 1
            .Width = 64 * 256
            .Height = 64 * 256
            .MoveUp = True
            .MoveDown = True
            .MoveLeft = True
            .MoveRight = True
            .XRef = 0
            .YRef = 0
            .Looping = False
            .Visible = False
            .SubMapWidth = 256
            .SubMapHeight = 256
            .ZOrder = 0
            .Tx1 = 0
            .Tx2 = 1
            .Ty1 = 0
            .Ty2 = 1
        End With
    For A = 1 To 64
    For B = 1 To 64
            With SubMap(Index, A, B)
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
End Sub

Public Sub QMResetAllMaps()
Dim T As Integer, A As Integer, B As Integer
For T = 1 To 8
        With Map(T)
            .XCount = 64
            .YCount = 64
            .XPos = 0
            .YPos = 0
            .XInc = 1
            .YInc = 1
            .Width = 64 * 256
            .Height = 64 * 256
            .MoveUp = True
            .MoveDown = True
            .MoveLeft = True
            .MoveRight = True
            .XRef = 0
            .YRef = 0
            .Looping = False
            .Visible = False
            .SubMapWidth = 256
            .SubMapHeight = 256
            .ZOrder = 0
            .Tx1 = 0
            .Tx2 = 1
            .Ty1 = 0
            .Ty2 = 1
        End With
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

Private Sub Class_Initialize()
Dim T As Integer
    For T = 1 To 8
        With Map(T)
            .XCount = 64
            .YCount = 64
            .XPos = 0
            .YPos = 0
            .XInc = 1
            .YInc = 1
            .Width = 64 * 256
            .Height = 64 * 256
            .MoveUp = True
            .MoveDown = True
            .MoveLeft = True
            .MoveRight = True
            .XRef = 0
            .YRef = 0
            .Looping = False
            .Visible = False
            .SubMapWidth = 256
            .SubMapHeight = 256
            .ZOrder = 0
            .Tx1 = 0
            .Tx2 = 1
            .Ty1 = 0
            .Ty2 = 1
        End With
    Next T
End Sub

Private Function TextureValue(Total As Single, Current As Single) As Single
Dim PixelTex As Single
PixelTex = 1 / (Total)
TextureValue = PixelTex * (Current)
End Function

