VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Add()
    OverlayCount = OverlayCount + 1
    ReDim Preserve Overlay(OverlayCount) As OverlayInfo
    
    With Overlay(OverlayCount)
        .UL.CRed = 255
        .UL.CGreen = 255
        .UL.CBlue = 255
        .UL.Alpha = 255
        .UR.CRed = 255
        .UR.CGreen = 255
        .UR.CBlue = 255
        .UR.Alpha = 255
        .LL.CRed = 255
        .LL.CGreen = 255
        .LL.CBlue = 255
        .LL.Alpha = 255
        .LR.CRed = 255
        .LR.CGreen = 255
        .LR.CBlue = 255
        .LR.Alpha = 255
        .Get.X = 0
        .Get.Y = 0
        .Get.Width = 16
        .Get.Height = 16
        .TextureIndex = 0
        .Visible = True
        .ZOrder = 0
        .RotationAngle = 0
        .XPos = 0
        .YPos = 0
        .Tx1 = 0
        .Tx2 = 1
        .Ty1 = 0
        .Ty2 = 1
        .C1 = D3DColorARGB(255, 255, 255, 255)
        .C2 = .C1
        .C3 = .C1
        .C4 = .C1
        .S1 = D3DColorARGB(255, 0, 0, 0)
        .S2 = .S1
        .S3 = .S1
        .S4 = .S1
        .XMirror = False
        .YMirror = False
    End With
    
End Sub

Public Sub AddFromSource(SourceIndex As Single)
    
    OverlayCount = OverlayCount + 1
    ReDim Preserve Overlay(OverlayCount) As OverlayInfo
    
    Overlay(OverlayCount) = Overlay(SourceIndex)
    
End Sub

Public Sub Swap(FromIndex As Integer, ToIndex As Integer)
    
    Overlay(0) = Overlay(FromIndex)
    Overlay(FromIndex) = Overlay(ToIndex)
    Overlay(ToIndex) = Overlay(0)
    
End Sub

Public Sub Copy(FromIndex As Integer, ToIndex As Integer)
    
    Overlay(ToIndex) = Overlay(FromIndex)
    
End Sub

Public Sub Clear()
    ReDim Overlay(0) As OverlayInfo
    OverlayCount = 0
End Sub

Public Sub SetTextureReference(Index As Integer, Value As Integer)
    Overlay(Index).TextureIndex = Value
End Sub
Public Sub SetVisible(Index As Integer, Value As Boolean)
    Overlay(Index).Visible = Value
End Sub
Public Sub SetRotationAngle(Index As Integer, Value As Single)
    Overlay(Index).RotationAngle = Value
End Sub
Public Sub SetZOrder(Index As Integer, Value As Integer)
    Overlay(Index).ZOrder = Value
End Sub
Public Sub SetXPosition(Index As Integer, Value As Single)
    Overlay(Index).XPos = Value
End Sub
Public Sub SetYPosition(Index As Integer, Value As Single)
    Overlay(Index).YPos = Value
End Sub

Public Sub SetXMirror(Index As Integer, Value As Boolean)
    Overlay(Index).XMirror = Value
    UpDateTValues Index
End Sub
Public Sub SetYMirror(Index As Integer, Value As Boolean)
    Overlay(Index).YMirror = Value
    UpDateTValues Index
End Sub
Public Function GetXMirror(Index As Integer) As Boolean
    GetXMirror = Overlay(Index).XMirror
End Function
Public Function GetYMirror(Index As Integer) As Boolean
    GetYMirror = Overlay(Index).YMirror
End Function

Public Sub SetULX(Index As Integer, Value As Single)
    Overlay(Index).UL.X = Value
End Sub
Public Sub SetULY(Index As Integer, Value As Single)
    Overlay(Index).UL.Y = Value
End Sub
Public Sub SetULCRed(Index As Integer, Value As Integer)
    Overlay(Index).UL.CRed = Value
    UpdateColor Index
End Sub
Public Sub SetULCGreen(Index As Integer, Value As Integer)
    Overlay(Index).UL.CGreen = Value
    UpdateColor Index
End Sub
Public Sub SetULCBlue(Index As Integer, Value As Integer)
    Overlay(Index).UL.CBlue = Value
    UpdateColor Index
End Sub
Public Sub SetULSRed(Index As Integer, Value As Integer)
    Overlay(Index).UL.SRed = Value
    UpdateSpecular Index
End Sub
Public Sub SetULSGreen(Index As Integer, Value As Integer)
    Overlay(Index).UL.SGreen = Value
    UpdateSpecular Index
End Sub
Public Sub SetULSBlue(Index As Integer, Value As Integer)
    Overlay(Index).UL.SBlue = Value
    UpdateSpecular Index
End Sub
Public Sub SetULAlpha(Index As Integer, Value As Integer)
    Overlay(Index).UL.Alpha = Value
    UpdateColor Index
End Sub
Public Sub ULSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .UL.CRed = Red
        .UL.CGreen = Green
        .UL.CBlue = Blue
    End With
    UpdateColor Index
End Sub
Public Sub ULSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .UL.SRed = Red
        .UL.SGreen = Green
        .UL.SBlue = Blue
    End With
    UpdateSpecular Index
End Sub

Public Sub SetURX(Index As Integer, Value As Single)
    Overlay(Index).UR.X = Value
End Sub
Public Sub SetURY(Index As Integer, Value As Single)
    Overlay(Index).UR.Y = Value
End Sub
Public Sub SetURCRed(Index As Integer, Value As Integer)
    Overlay(Index).UR.CRed = Value
    UpdateColor Index
End Sub
Public Sub SetURCGreen(Index As Integer, Value As Integer)
    Overlay(Index).UR.CGreen = Value
    UpdateColor Index
End Sub
Public Sub SetURCBlue(Index As Integer, Value As Integer)
    Overlay(Index).UR.CBlue = Value
    UpdateColor Index
End Sub
Public Sub SetURSRed(Index As Integer, Value As Integer)
    Overlay(Index).UR.SRed = Value
    UpdateSpecular Index
End Sub
Public Sub SetURSGreen(Index As Integer, Value As Integer)
    Overlay(Index).UR.SGreen = Value
    UpdateSpecular Index
End Sub
Public Sub SetURSBlue(Index As Integer, Value As Integer)
    Overlay(Index).UR.SBlue = Value
    UpdateSpecular Index
End Sub
Public Sub SetURAlpha(Index As Integer, Value As Integer)
    Overlay(Index).UR.Alpha = Value
    UpdateColor Index
End Sub
Public Sub URSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .UR.CRed = Red
        .UR.CGreen = Green
        .UR.CBlue = Blue
    End With
    UpdateColor Index
End Sub
Public Sub URSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .UR.SRed = Red
        .UR.SGreen = Green
        .UR.SBlue = Blue
    End With
    UpdateSpecular Index
End Sub

Public Sub SetLLX(Index As Integer, Value As Single)
    Overlay(Index).LL.X = Value
End Sub
Public Sub SetLLY(Index As Integer, Value As Single)
    Overlay(Index).LL.Y = Value
End Sub
Public Sub SetLLCRed(Index As Integer, Value As Integer)
    Overlay(Index).LL.CRed = Value
    UpdateColor Index
End Sub
Public Sub SetLLCGreen(Index As Integer, Value As Integer)
    Overlay(Index).LL.CGreen = Value
    UpdateColor Index
End Sub
Public Sub SetLLCBlue(Index As Integer, Value As Integer)
    Overlay(Index).LL.CBlue = Value
    UpdateColor Index
End Sub
Public Sub SetLLSRed(Index As Integer, Value As Integer)
    Overlay(Index).LL.SRed = Value
    UpdateSpecular Index
End Sub
Public Sub SetLLSGreen(Index As Integer, Value As Integer)
    Overlay(Index).LL.SGreen = Value
    UpdateSpecular Index
End Sub
Public Sub SetLLSBlue(Index As Integer, Value As Integer)
    Overlay(Index).LL.SBlue = Value
    UpdateSpecular Index
End Sub
Public Sub SetLLAlpha(Index As Integer, Value As Integer)
    Overlay(Index).LL.Alpha = Value
    UpdateColor Index
End Sub
Public Sub LLSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .LL.CRed = Red
        .LL.CGreen = Green
        .LL.CBlue = Blue
    End With
    UpdateColor Index
End Sub
Public Sub LLSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .LL.SRed = Red
        .LL.SGreen = Green
        .LL.SBlue = Blue
    End With
    UpdateSpecular Index
End Sub

Public Sub SetLRX(Index As Integer, Value As Single)
    Overlay(Index).LR.X = Value
End Sub
Public Sub SetLRY(Index As Integer, Value As Single)
    Overlay(Index).LR.Y = Value
End Sub
Public Sub SetLRCRed(Index As Integer, Value As Integer)
    Overlay(Index).LR.CRed = Value
    UpdateColor Index
End Sub
Public Sub SetLRCGreen(Index As Integer, Value As Integer)
    Overlay(Index).LR.CGreen = Value
    UpdateColor Index
End Sub
Public Sub SetLRCBlue(Index As Integer, Value As Integer)
    Overlay(Index).LR.CBlue = Value
    UpdateColor Index
End Sub
Public Sub SetLRSRed(Index As Integer, Value As Integer)
    Overlay(Index).LR.SRed = Value
    UpdateSpecular Index
End Sub
Public Sub SetLRSGreen(Index As Integer, Value As Integer)
    Overlay(Index).LR.SGreen = Value
    UpdateSpecular Index
End Sub
Public Sub SetLRSBlue(Index As Integer, Value As Integer)
    Overlay(Index).LR.SBlue = Value
    UpdateSpecular Index
End Sub
Public Sub SetLRAlpha(Index As Integer, Value As Integer)
    Overlay(Index).LR.Alpha = Value
    UpdateColor Index
End Sub
Public Sub LRSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .LR.CRed = Red
        .LR.CGreen = Green
        .LR.CBlue = Blue
    End With
    UpdateColor Index
End Sub
Public Sub LRSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .LR.SRed = Red
        .LR.SGreen = Green
        .LR.SBlue = Blue
    End With
    UpdateSpecular Index
End Sub

Public Sub SetGetX(Index As Integer, Value As Single)
    Overlay(Index).Get.X = Value
    UpDateTValues Index
End Sub
Public Sub SetGetY(Index As Integer, Value As Single)
    Overlay(Index).Get.Y = Value
    UpDateTValues Index
End Sub
Public Sub SetGetWidth(Index As Integer, Value As Single)
    Overlay(Index).Get.Width = Value
    UpDateTValues Index
End Sub
Public Sub SetGetHeight(Index As Integer, Value As Single)
    Overlay(Index).Get.Height = Value
    UpDateTValues Index
End Sub

Public Function GetTextureReference(Index As Integer) As Integer
    GetTextureReference = Overlay(Index).TextureIndex
End Function
Public Function GetVisible(Index As Integer) As Boolean
    GetVisible = Overlay(Index).Visible
End Function
Public Function GetRotationAngle(Index As Integer) As Single
    GetRotationAngle = Overlay(Index).RotationAngle
End Function
Public Function GetZOrder(Index As Integer) As Integer
    GetZOrder = Overlay(Index).ZOrder
End Function
Public Function GetXPosition(Index As Integer) As Single
    GetXPosition = Overlay(Index).XPos
End Function
Public Function GetYPosition(Index As Integer) As Single
    GetYPosition = Overlay(Index).YPos
End Function

Public Function GetULX(Index As Integer) As Single
    GetULX = Overlay(Index).UL.X
End Function
Public Function GetULY(Index As Integer) As Single
    GetULY = Overlay(Index).UL.Y
End Function
Public Function GetULCRed(Index As Integer) As Integer
    GetULCRed = Overlay(Index).UL.CRed
End Function
Public Function GetULCGreen(Index As Integer) As Integer
    GetULCGreen = Overlay(Index).UL.CGreen
End Function
Public Function GetULCBlue(Index As Integer) As Integer
    GetULCBlue = Overlay(Index).UL.CBlue
End Function
Public Function GetULSRed(Index As Integer) As Integer
    GetULSRed = Overlay(Index).UL.SRed
End Function
Public Function GetULSGreen(Index As Integer) As Integer
    GetULSGreen = Overlay(Index).UL.SGreen
End Function
Public Function GetULSBlue(Index As Integer) As Integer
    GetULSBlue = Overlay(Index).UL.SBlue
End Function
Public Function GetULAlpha(Index As Integer) As Integer
    GetULAlpha = Overlay(Index).UL.Alpha
End Function

Public Function GetURX(Index As Integer) As Single
    GetURX = Overlay(Index).UR.X
End Function
Public Function GetURY(Index As Integer) As Single
    GetURY = Overlay(Index).UR.Y
End Function
Public Function GetURCRed(Index As Integer) As Integer
    GetURCRed = Overlay(Index).UR.CRed
End Function
Public Function GetURCGreen(Index As Integer) As Integer
    GetURCGreen = Overlay(Index).UR.CGreen
End Function
Public Function GetURCBlue(Index As Integer) As Integer
    GetURCBlue = Overlay(Index).UR.CBlue
End Function
Public Function GetURSRed(Index As Integer) As Integer
    GetURSRed = Overlay(Index).UR.SRed
End Function
Public Function GetURSGreen(Index As Integer) As Integer
    GetURSGreen = Overlay(Index).UR.SGreen
End Function
Public Function GetURSBlue(Index As Integer) As Integer
    GetURSBlue = Overlay(Index).UR.SBlue
End Function
Public Function GetURAlpha(Index As Integer) As Integer
    GetURAlpha = Overlay(Index).UR.Alpha
End Function

Public Function GetLLX(Index As Integer) As Single
    GetLLX = Overlay(Index).LL.X
End Function
Public Function GetLLY(Index As Integer) As Single
    GetLLY = Overlay(Index).LL.Y
End Function
Public Function GetLLCRed(Index As Integer) As Integer
    GetLLCRed = Overlay(Index).LL.CRed
End Function
Public Function GetLLCGreen(Index As Integer) As Integer
    GetLLCGreen = Overlay(Index).LL.CGreen
End Function
Public Function GetLLCBlue(Index As Integer) As Integer
    GetLLCBlue = Overlay(Index).LL.CBlue
End Function
Public Function GetLLSRed(Index As Integer) As Integer
    GetLLSRed = Overlay(Index).LL.SRed
End Function
Public Function GetLLSGreen(Index As Integer) As Integer
    GetLLSGreen = Overlay(Index).LL.SGreen
End Function
Public Function GetLLSBlue(Index As Integer) As Integer
    GetLLSBlue = Overlay(Index).LL.SBlue
End Function
Public Function GetLLAlpha(Index As Integer) As Integer
    GetLLAlpha = Overlay(Index).LL.Alpha
End Function

Public Function GetLRX(Index As Integer) As Single
    GetLRX = Overlay(Index).LR.X
End Function
Public Function GetLRY(Index As Integer) As Single
    GetLRY = Overlay(Index).LR.Y
End Function
Public Function GetLRCRed(Index As Integer) As Integer
    GetLRCRed = Overlay(Index).LR.CRed
End Function
Public Function GetLRCGreen(Index As Integer) As Integer
    GetLRCGreen = Overlay(Index).LR.CGreen
End Function
Public Function GetLRCBlue(Index As Integer) As Integer
    GetLRCBlue = Overlay(Index).LR.CBlue
End Function
Public Function GetLRSRed(Index As Integer) As Integer
    GetLRSRed = Overlay(Index).LR.SRed
End Function
Public Function GetLRSGreen(Index As Integer) As Integer
    GetLRSGreen = Overlay(Index).LR.SGreen
End Function
Public Function GetLRSBlue(Index As Integer) As Integer
    GetLRSBlue = Overlay(Index).LR.SBlue
End Function
Public Function GetLRAlpha(Index As Integer) As Integer
    GetLRAlpha = Overlay(Index).LR.Alpha
End Function

Public Function GetGetX(Index As Integer) As Single
    GetGetX = Overlay(Index).Get.X
End Function
Public Function GetGetY(Index As Integer) As Single
    GetGetY = Overlay(Index).Get.Y
End Function
Public Function GetGetWidth(Index As Integer) As Single
    GetGetWidth = Overlay(Index).Get.Width
End Function
Public Function GetGetHeight(Index As Integer) As Single
    GetGetHeight = Overlay(Index).Get.Height
End Function

Public Sub QMSetPutRectangle(Index As Integer, X As Single, Y As Single, Width As Single, Height As Single)
    With Overlay(Index).UL
        .X = X
        .Y = Y
    End With
    With Overlay(Index).UR
        .X = X + Width
        .Y = Y
    End With
    With Overlay(Index).LL
        .X = X
        .Y = Y + Height
    End With
    With Overlay(Index).LR
        .X = X + Width
        .Y = Y + Height
    End With
End Sub

Public Sub QMSetGetRectangle(Index As Integer, X As Single, Y As Single, Width As Single, Height As Single)
    Overlay(Index).Get.X = X
    Overlay(Index).Get.Y = Y
    Overlay(Index).Get.Width = Width
    Overlay(Index).Get.Height = Height
    UpDateTValues Index
End Sub

Public Sub QMSetCRed(Index As Integer, Value As Integer)
    Overlay(Index).UL.CRed = Value
    Overlay(Index).UR.CRed = Value
    Overlay(Index).LL.CRed = Value
    Overlay(Index).LR.CRed = Value
    UpdateColor Index
End Sub

Public Sub QMSetCGreen(Index As Integer, Value As Integer)
    Overlay(Index).UL.CGreen = Value
    Overlay(Index).UR.CGreen = Value
    Overlay(Index).LL.CGreen = Value
    Overlay(Index).LR.CGreen = Value
    UpdateColor Index
End Sub
Public Sub QMSetCBlue(Index As Integer, Value As Integer)
    Overlay(Index).UL.CBlue = Value
    Overlay(Index).UR.CBlue = Value
    Overlay(Index).LL.CBlue = Value
    Overlay(Index).LR.CBlue = Value
    UpdateColor Index
End Sub

Public Sub QMSetSRed(Index As Integer, Value As Integer)
    Overlay(Index).UL.SRed = Value
    Overlay(Index).UR.SRed = Value
    Overlay(Index).LL.SRed = Value
    Overlay(Index).LR.SRed = Value
    UpdateSpecular Index
End Sub
Public Sub QMSetSGreen(Index As Integer, Value As Integer)
    Overlay(Index).UL.SGreen = Value
    Overlay(Index).UR.SGreen = Value
    Overlay(Index).LL.SGreen = Value
    Overlay(Index).LR.SGreen = Value
    UpdateSpecular Index
End Sub
Public Sub QMSetSBlue(Index As Integer, Value As Integer)
    Overlay(Index).UL.SBlue = Value
    Overlay(Index).UR.SBlue = Value
    Overlay(Index).LL.SBlue = Value
    Overlay(Index).LR.SBlue = Value
    UpdateSpecular Index
End Sub

Public Sub QMSetAlpha(Index As Integer, Value As Integer)
    Overlay(Index).UL.Alpha = Value
    Overlay(Index).UR.Alpha = Value
    Overlay(Index).LL.Alpha = Value
    Overlay(Index).LR.Alpha = Value
    UpdateColor Index
End Sub

Public Sub QMSetAllPoints(Index As Integer, ULX As Single, ULY As Single, URX As Single, URY As Single, LLX As Single, LLY As Single, LRX As Single, LRY As Single)
    With Overlay(Index)
        .UL.X = ULX
        .UL.Y = ULY
        .UR.X = URX
        .UR.Y = URY
        .LL.X = LLX
        .LL.Y = LLY
        .LR.X = LRX
        .LR.Y = LRY
    End With
End Sub

Public Sub QMSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .UL.CRed = Red
        .UL.CGreen = Green
        .UL.CBlue = Blue
        .UR.CRed = Red
        .UR.CGreen = Green
        .UR.CBlue = Blue
        .LL.CRed = Red
        .LL.CGreen = Green
        .LL.CBlue = Blue
        .LR.CRed = Red
        .LR.CGreen = Green
        .LR.CBlue = Blue
    End With
    UpdateColor Index
End Sub

Public Sub QMSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Overlay(Index)
        .UL.SRed = Red
        .UR.SRed = Red
        .LL.SRed = Red
        .LR.SRed = Red
        .UL.SGreen = Green
        .UR.SGreen = Green
        .LL.SGreen = Green
        .LR.SGreen = Green
        .UL.SBlue = Blue
        .UR.SBlue = Blue
        .LL.SBlue = Blue
        .LR.SBlue = Blue
    End With
    UpdateSpecular Index
End Sub

Public Sub QMExpand(Index As Integer, Value As Single)
    With Overlay(Index)
        .UL.X = .UL.X - Value
        .UL.Y = .UL.Y - Value
        .UR.X = .UR.X + Value
        .UR.Y = .UR.Y - Value
        .LL.X = .LL.X - Value
        .LL.Y = .LL.Y + Value
        .LR.X = .LR.X + Value
        .LR.Y = .LR.Y + Value
    End With
End Sub

Public Sub QMExpandWidth(Index As Integer, Value As Single)
    With Overlay(Index)
        .UL.X = .UL.X - Value
        .UR.X = .UR.X + Value
        .LL.X = .LL.X - Value
        .LR.X = .LR.X + Value
    End With
End Sub

Public Sub QMEXpandHeight(Index As Integer, Value As Single)
    With Overlay(Index)
        .UL.Y = .UL.Y - Value
        .UR.Y = .UR.Y - Value
        .LL.Y = .LL.Y + Value
        .LR.Y = .LR.Y + Value
    End With
End Sub

Public Function GetOverlayCount() As Integer
    GetOverlayCount = OverlayCount
End Function

Private Sub Class_Initialize()
    ReDim Overlay(0) As OverlayInfo
End Sub

Private Sub Class_Terminate()
    ReDim Overlay(0) As OverlayInfo
End Sub

Private Sub UpdateColor(Index As Integer)
    With Overlay(Index)
        .C1 = D3DColorARGB(.UL.Alpha, .UL.CRed, .UL.CGreen, .UL.CBlue)
        .C2 = D3DColorARGB(.UR.Alpha, .UR.CRed, .UR.CGreen, .UR.CBlue)
        .C3 = D3DColorARGB(.LL.Alpha, .LL.CRed, .LL.CGreen, .LL.CBlue)
        .C4 = D3DColorARGB(.LR.Alpha, .LR.CRed, .LR.CGreen, .LR.CBlue)
    End With
End Sub

Private Sub UpdateSpecular(Index As Integer)
    With Overlay(Index)
        .S1 = D3DColorARGB(255, .UL.SRed, .UL.SGreen, .UL.SBlue)
        .S2 = D3DColorARGB(255, .UR.SRed, .UR.SGreen, .UR.SBlue)
        .S3 = D3DColorARGB(255, .LL.SRed, .LL.SGreen, .LL.SBlue)
        .S4 = D3DColorARGB(255, .LR.SRed, .LR.SGreen, .LR.SBlue)
    End With
End Sub

Private Sub UpDateTValues(Index As Integer)
Dim TSize As Single
If Overlay(Index).TextureIndex = 0 Then Exit Sub
    TSize = TextureInfo(Overlay(Index).TextureIndex).Size
    With Overlay(Index)
    If .XMirror = False Then
        .Tx1 = TextureValue(TSize, .Get.X)
        .Tx2 = TextureValue(TSize, .Get.X + .Get.Width)
    Else:
        .Tx1 = TextureValue(TSize, .Get.X + .Get.Width)
        .Tx2 = TextureValue(TSize, .Get.X)
    End If
    If .YMirror = False Then
        .Ty1 = TextureValue(TSize, .Get.Y)
        .Ty2 = TextureValue(TSize, .Get.Y + .Get.Height)
    Else:
        .Ty1 = TextureValue(TSize, .Get.Y + .Get.Height)
        .Ty2 = TextureValue(TSize, .Get.Y)
    End If
    End With
End Sub

Private Function TextureValue(Total As Single, Current As Single) As Single
Dim PixelTex As Single
PixelTex = 1 / (Total)
TextureValue = PixelTex * (Current)
End Function

