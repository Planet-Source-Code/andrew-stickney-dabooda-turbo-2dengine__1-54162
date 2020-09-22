VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBMapData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MapData() As Integer

Private Width As Single
Private Height As Single
Private TWidth As Single, THeight As Single
Private XRemainder As Single, YRemainder As Single

Public Sub SetUpMapDataFromMap(MapIndex As Integer, TileWidth As Integer, TileHeight As Integer)
Dim W As Single, H As Single
W = Map(MapIndex).SubMapWidth
H = Map(MapIndex).SubMapHeight
Dim xc As Integer, yc As Integer
xc = Map(MapIndex).XCount
yc = Map(MapIndex).YCount
Dim A As Integer, B As Integer

Width = xc * W
Height = yc * H

TWidth = TileWidth
THeight = TileHeight

A = Int(Width / TWidth) - 1
B = Int(Height / THeight) - 1

ReDim MapData(A, B) As Integer

End Sub

Public Sub SetUpMapData(XCount As Integer, YCount As Integer, TileWidth As Integer, TileHeight As Integer)
    Width = XCount * TileWidth
    Height = YCount * TileHeight
    TWidth = TileWidth
    THeight = TileHeight
    XCount = XCount - 1
    YCount = YCount - 1
    ReDim MapData(XCount, YCount) As Integer
End Sub

Public Sub SetMapData(XIndex As Integer, YIndex As Integer, Data As Integer)
    MapData(XIndex - 1, YIndex - 1) = Data
End Sub

Public Sub DeStroy()
    ReDim MapData(0) As Integer
End Sub

Public Sub SetAll(Data As Integer)
    Dim xc As Integer, yc As Integer
    Dim X As Integer, Y As Integer
    X = Int(Width / TWidth) - 1
    Y = Int(Height / THeight) - 1
    
    For xc = 0 To X
    For yc = 0 To Y
        MapData(xc, yc) = Data
    Next yc, xc
    
End Sub

Public Function GetMapData(XIndex As Integer, YIndex As Integer) As Integer
    GetMapData = MapData(XIndex - 1, YIndex - 1)
End Function

Public Function GetMapDataByPosition(X As Single, Y As Single) As Integer
    Dim xc As Integer, yc As Integer
    If X > Width - 1 Then X = X - Width
    If X < 0 Then X = Width + X
    If Y > Height - 1 Then Y = Y - Height
    If Y < 0 Then Y = Height + Y
    
    xc = Int(X / TWidth)
    yc = Int(Y / THeight)
    
    XRemainder = X - (xc * TWidth)
    YRemainder = Y - (yc * THeight)
    
    GetMapDataByPosition = MapData(xc, yc)
End Function

Public Function GetMapDataBySprite(SpriteIndex As Integer, XOffset As Single, YOffset As Single) As Integer
    Dim X As Single, Y As Single
    Dim xc As Integer, yc As Integer
    
    X = Sprite(SpriteIndex).XPos + XOffset
    Y = Sprite(SpriteIndex).YPos + YOffset
    
    If X > Width - 1 Then X = X - Width
    If X < 0 Then X = Width + X
    If Y > Height - 1 Then Y = Y - Height
    If Y < 0 Then Y = Height + Y
    
    xc = Int(X / TWidth)
    yc = Int(Y / THeight)
    
    XRemainder = X - (xc * TWidth)
    YRemainder = Y - (yc * THeight)
    
    GetMapDataBySprite = MapData(xc, yc)
End Function

Public Function GetXCount() As Integer
    GetXCount = Int(Width / TWidth)
End Function
Public Function GetYCount() As Integer
    GetYCount = Int(Height / THeight)
End Function

Public Function GetTileWidth() As Integer
    GetTileWidth = TWidth
End Function
Public Function GetTileHeight() As Integer
    GetTileHeight = THeight
End Function

Public Function GetXRemainder() As Integer
    GetXRemainder = XRemainder
End Function
Public Function GetYRemainder() As Integer
    GetYRemainder = YRemainder
End Function

Private Sub Class_Initialize()
ReDim MapData(0) As Integer
End Sub

Private Sub Class_Terminate()
ReDim MapData(0) As Integer
End Sub
