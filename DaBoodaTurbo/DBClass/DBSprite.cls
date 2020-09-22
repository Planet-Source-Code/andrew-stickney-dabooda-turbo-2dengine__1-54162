VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Add()
    SpriteCount = SpriteCount + 1
    ReDim Preserve Sprite(SpriteCount) As SpriteInfo
    
    With Sprite(SpriteCount)
        .XPos = 0
        .YPos = 0
        .Radius = 8
        .RAngle = 0
        .TRef = 0
        .Sref = 0
        .Mref = 0
        With .Get
            .X = 0
            .Y = 0
            .Width = 16
            .Height = 16
        End With
        With .UL
            .X = -8
            .Y = -8
            .Alpha = 255
            .CRed = 255
            .CGreen = 255
            .CBlue = 255
            .SRed = 0
            .SGreen = 0
            .SBlue = 0
        End With
        With .UR
            .X = -8
            .Y = -8
            .Alpha = 255
            .CRed = 255
            .CGreen = 255
            .CBlue = 255
            .SRed = 0
            .SGreen = 0
            .SBlue = 0
        End With
        With .LL
            .X = -8
            .Y = -8
            .Alpha = 255
            .CRed = 255
            .CGreen = 255
            .CBlue = 255
            .SRed = 0
            .SGreen = 0
            .SBlue = 0
        End With
        With .LR
            .X = -8
            .Y = -8
            .Alpha = 255
            .CRed = 255
            .CGreen = 255
            .CBlue = 255
            .SRed = 0
            .SGreen = 0
            .SBlue = 0
        End With
        .ZOrder = 0
        .Visible = True
        .Drawn = False
        .DrawnX = 0
        .DrawnY = 0
        .Counter = 0
        .CounterInc = 1
        .AutoCounterInc = False
        .Tx1 = 0
        .Tx2 = 1
        .Ty1 = 0
        .Ty2 = 1
        .C1 = D3DColorARGB(255, 255, 255, 255)
        .C2 = D3DColorARGB(255, 255, 255, 255)
        .C3 = D3DColorARGB(255, 255, 255, 255)
        .C4 = D3DColorARGB(255, 255, 255, 255)
        .S1 = D3DColorARGB(255, 0, 0, 0)
        .S2 = D3DColorARGB(255, 0, 0, 0)
        .S3 = D3DColorARGB(255, 0, 0, 0)
        .S4 = D3DColorARGB(255, 0, 0, 0)
        .Void = False
        .XMirror = False
        .YMirror = False
        .InheritAngle = True
    End With
    
End Sub

Public Sub AddFromSource(SourceIndex As Integer)
    SpriteCount = SpriteCount + 1
    ReDim Preserve Sprite(SpriteCount) As SpriteInfo
    
    Sprite(SpriteCount) = Sprite(SourceIndex)
End Sub

Public Sub Swap(FromIndex As Integer, ToIndex As Integer)
    Sprite(0) = Sprite(FromIndex)
    Sprite(FromIndex) = Sprite(ToIndex)
    Sprite(ToIndex) = Sprite(0)
End Sub

Public Sub Copy(FromIndex As Integer, ToIndex As Integer)
    Sprite(ToIndex) = Sprite(FromIndex)
End Sub

Public Sub Clear()
    ReDim Sprite(0) As SpriteInfo
    SpriteCount = 0
End Sub

Public Sub RemoveSprite(Index As Integer)
    Dim T As Integer
    If Index <> SpriteCount Then
        For T = Index To SpriteCount - 1
            Sprite(T) = Sprite(T + 1)
        Next T
        SpriteCount = SpriteCount - 1
        ReDim Preserve Sprite(SpriteCount) As SpriteInfo
    Else:
        SpriteCount = SpriteCount - 1
        ReDim Preserve Sprite(SpriteCount) As SpriteInfo
    End If
End Sub

Public Sub Void(Index As Integer)
    Sprite(Index).Void = True
End Sub
Public Sub UnVoid(Index As Integer)
    Sprite(Index).Void = False
End Sub

Public Sub SetXMirror(Index As Integer, Value As Boolean)
    Sprite(Index).XMirror = Value
    UpDateTValues Index
End Sub
Public Sub SetYMirror(Index As Integer, Value As Boolean)
    Sprite(Index).YMirror = Value
    UpDateTValues Index
End Sub
Public Function GetXMirror(Index As Integer) As Boolean
    GetXMirror = Sprite(Index).XMirror
End Function
Public Function GetYMirror(Index As Integer) As Boolean
    GetYMirror = Sprite(Index).YMirror
End Function

Public Sub SetInheritAngle(Index As Integer, Value As Boolean)
    Sprite(Index).InheritAngle = Value
End Sub
Public Function GetInheritAngle(Index As Integer) As Boolean
    GetInheritAngle = Sprite(Index).InheritAngle
End Function

Public Function GetSpriteCount() As Integer
    GetSpriteCount = SpriteCount
End Function

Public Sub SetXPosition(Index As Integer, Value As Single)
    Sprite(Index).XPos = Value
End Sub
Public Sub SetYPosition(Index As Integer, Value As Single)
    Sprite(Index).YPos = Value
End Sub
Public Sub SetRadius(Index As Integer, Value As Single)
    Sprite(Index).Radius = Value
End Sub
Public Sub SetRotationAngle(Index As Integer, Value As Single)
    Sprite(Index).RAngle = Value
End Sub
Public Sub SetTextureReference(Index As Integer, Value As Integer)
    Sprite(Index).TRef = Value
End Sub
Public Sub SetSpriteReference(Index As Integer, Value As Integer)
    Sprite(Index).Sref = Value
End Sub
Public Sub SetMapReference(Index As Integer, Value As Integer)
    Sprite(Index).Mref = Value
End Sub

Public Sub GetSetX(Index As Integer, Value As Single)
    Sprite(Index).Get.X = Value
    UpDateTValues Index
End Sub
Public Sub GetSetY(Index As Integer, Value As Single)
    Sprite(Index).Get.Y = Value
    UpDateTValues Index
End Sub
Public Sub GetSetWidth(Index As Integer, Value As Single)
    Sprite(Index).Get.Width = Value
    UpDateTValues Index
End Sub
Public Sub GetSetHeight(Index As Integer, Value As Single)
    Sprite(Index).Get.Height = Value
    UpDateTValues Index
End Sub

Public Sub ULSetX(Index As Integer, Value As Single)
    Sprite(Index).UL.X = Value
End Sub
Public Sub ULSetY(Index As Integer, Value As Single)
    Sprite(Index).UL.Y = Value
End Sub
Public Sub ULSetAlpha(Index As Integer, Value As Integer)
    Sprite(Index).UL.Alpha = Value
    UpdateColor Index
End Sub
Public Sub ULSetCRed(Index As Integer, Value As Integer)
    Sprite(Index).UL.CRed = Value
    UpdateColor Index
End Sub
Public Sub ULSetCGreen(Index As Integer, Value As Integer)
    Sprite(Index).UL.CGreen = Value
    UpdateColor Index
End Sub
Public Sub ULSetCBlue(Index As Integer, Value As Integer)
    Sprite(Index).UL.CBlue = Value
    UpdateColor Index
End Sub
Public Sub ULSetSRed(Index As Integer, Value As Integer)
    Sprite(Index).UL.SRed = Value
    UpdateSpecular Index
End Sub
Public Sub ULSetSGreen(Index As Integer, Value As Integer)
    Sprite(Index).UL.SGreen = Value
    UpdateSpecular Index
End Sub
Public Sub ULSetSBlue(Index As Integer, Value As Integer)
    Sprite(Index).UL.SBlue = Value
    UpdateSpecular Index
End Sub
Public Sub ULSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        .UL.CRed = Red
        .UL.CGreen = Green
        .UL.CBlue = Blue
    End With
    UpdateColor Index
End Sub
Public Sub ULSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        .UL.SRed = Red
        .UL.SGreen = Green
        .UL.SBlue = Blue
    End With
    UpdateSpecular Index
End Sub

Public Sub URSetX(Index As Integer, Value As Single)
    Sprite(Index).UR.X = Value
End Sub
Public Sub URSetY(Index As Integer, Value As Single)
    Sprite(Index).UR.Y = Value
End Sub
Public Sub URSetAlpha(Index As Integer, Value As Integer)
    Sprite(Index).UR.Alpha = Value
    UpdateColor Index
End Sub
Public Sub URSetCRed(Index As Integer, Value As Integer)
    Sprite(Index).UR.CRed = Value
    UpdateColor Index
End Sub
Public Sub URSetCGreen(Index As Integer, Value As Integer)
    Sprite(Index).UR.CGreen = Value
    UpdateColor Index
End Sub
Public Sub URSetCBlue(Index As Integer, Value As Integer)
    Sprite(Index).UR.CBlue = Value
    UpdateColor Index
End Sub
Public Sub URSetSRed(Index As Integer, Value As Integer)
    Sprite(Index).UR.SRed = Value
    UpdateSpecular Index
End Sub
Public Sub URSetSGreen(Index As Integer, Value As Integer)
    Sprite(Index).UR.SGreen = Value
    UpdateSpecular Index
End Sub
Public Sub URSetSBlue(Index As Integer, Value As Integer)
    Sprite(Index).UR.SBlue = Value
    UpdateSpecular Index
End Sub
Public Sub URSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        .UR.CRed = Red
        .UR.CGreen = Green
        .UR.CBlue = Blue
    End With
    UpdateColor Index
End Sub
Public Sub URSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        .UR.SRed = Red
        .UR.SGreen = Green
        .UR.SBlue = Blue
    End With
    UpdateSpecular Index
End Sub

Public Sub LLSetX(Index As Integer, Value As Single)
    Sprite(Index).LL.X = Value
End Sub
Public Sub LLSetY(Index As Integer, Value As Single)
    Sprite(Index).LL.Y = Value
End Sub
Public Sub LLSetAlpha(Index As Integer, Value As Integer)
    Sprite(Index).LL.Alpha = Value
    UpdateColor Index
End Sub
Public Sub LLSetCRed(Index As Integer, Value As Integer)
    Sprite(Index).LL.CRed = Value
    UpdateColor Index
End Sub
Public Sub LLSetCGreen(Index As Integer, Value As Integer)
    Sprite(Index).LL.CGreen = Value
    UpdateColor Index
End Sub
Public Sub LLSetCBlue(Index As Integer, Value As Integer)
    Sprite(Index).LL.CBlue = Value
    UpdateColor Index
End Sub
Public Sub LLSetSRed(Index As Integer, Value As Integer)
    Sprite(Index).LL.SRed = Value
    UpdateSpecular Index
End Sub
Public Sub LLSetSGreen(Index As Integer, Value As Integer)
    Sprite(Index).LL.SGreen = Value
    UpdateSpecular Index
End Sub
Public Sub LLSetSBlue(Index As Integer, Value As Integer)
    Sprite(Index).LL.SBlue = Value
    UpdateSpecular Index
End Sub
Public Sub LLSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        .LL.CRed = Red
        .LL.CGreen = Green
        .LL.CBlue = Blue
    End With
    UpdateColor Index
End Sub
Public Sub LLSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        .LL.SRed = Red
        .LL.SGreen = Green
        .LL.SBlue = Blue
    End With
    UpdateSpecular Index
End Sub

Public Sub LRSetX(Index As Integer, Value As Single)
    Sprite(Index).LR.X = Value
End Sub
Public Sub LRSetY(Index As Integer, Value As Single)
    Sprite(Index).LR.Y = Value
End Sub
Public Sub LRSetAlpha(Index As Integer, Value As Integer)
    Sprite(Index).LR.Alpha = Value
    UpdateColor Index
End Sub
Public Sub LRSetCRed(Index As Integer, Value As Integer)
    Sprite(Index).LR.CRed = Value
    UpdateColor Index
End Sub
Public Sub LRSetCGreen(Index As Integer, Value As Integer)
    Sprite(Index).LR.CGreen = Value
    UpdateColor Index
End Sub
Public Sub LRSetCBlue(Index As Integer, Value As Integer)
    Sprite(Index).LR.CBlue = Value
    UpdateColor Index
End Sub
Public Sub LRSetSRed(Index As Integer, Value As Integer)
    Sprite(Index).LR.SRed = Value
    UpdateSpecular Index
End Sub
Public Sub LRSetSGreen(Index As Integer, Value As Integer)
    Sprite(Index).LR.SGreen = Value
    UpdateSpecular Index
End Sub
Public Sub LRSetSBlue(Index As Integer, Value As Integer)
    Sprite(Index).LR.SBlue = Value
    UpdateSpecular Index
End Sub
Public Sub LRSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        .LR.CRed = Red
        .LR.CGreen = Green
        .LR.CBlue = Blue
    End With
    UpdateColor Index
End Sub
Public Sub LRSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        .LR.SRed = Red
        .LR.SGreen = Green
        .LR.SBlue = Blue
    End With
    UpdateSpecular Index
End Sub

Public Sub SetZOrder(Index As Integer, Value As Integer)
    Sprite(Index).ZOrder = Value
End Sub
Public Sub SetVisible(Index As Integer, Value As Boolean)
    Sprite(Index).Visible = Value
End Sub

Public Sub CounterSet(Index As Integer, Value As Single)
    Sprite(Index).Counter = Value
End Sub
Public Sub CounterSetIncrement(Index As Integer, Value As Single)
    Sprite(Index).CounterInc = Value
End Sub
Public Sub CounterAuto(Index As Integer, Value As Boolean)
    Sprite(Index).AutoCounterInc = Value
End Sub

Public Function GetXPosition(Index As Integer) As Single
    GetXPosition = Sprite(Index).XPos
End Function
Public Function GetYPosition(Index As Integer) As Single
    GetYPosition = Sprite(Index).YPos
End Function
Public Function GetRadius(Index As Integer) As Single
    GetRadius = Sprite(Index).Radius
End Function
Public Function GetRotationAngle(Index As Integer) As Single
    GetRotationAngle = Sprite(Index).RAngle
End Function

Public Function GetTextureReference(Index As Integer) As Integer
    GetTextureReference = Sprite(Index).TRef
End Function
Public Function GetSpriteReference(Index As Integer) As Integer
    GetSpriteReference = Sprite(Index).Sref
End Function
Public Function GetMapReference(Index As Integer) As Integer
    GetMapReference = Sprite(Index).Mref
End Function

Public Function GetGetX(Index As Integer) As Single
    GetGetX = Sprite(Index).Get.X
End Function
Public Function GetGetY(Index As Integer) As Single
    GetGetY = Sprite(Index).Get.Y
End Function
Public Function GetGetWidth(Index As Integer) As Single
    GetGetWidth = Sprite(Index).Get.Width
End Function
Public Function GetGetHeight(Index As Integer) As Single
    GetGetHeight = Sprite(Index).Get.Height
End Function

Public Function ULGetX(Index As Integer) As Single
    ULGetX = Sprite(Index).UL.X
End Function
Public Function ULGetY(Index As Integer) As Single
    ULGetY = Sprite(Index).UL.Y
End Function
Public Function ULGetAlpha(Index As Integer) As Integer
    ULGetAlpha = Sprite(Index).UL.Alpha
End Function
Public Function ULGetCRed(Index As Integer) As Integer
    ULGetCRed = Sprite(Index).UL.CRed
End Function
Public Function ULGetCGreen(Index As Integer) As Integer
    ULGetCGreen = Sprite(Index).UL.CGreen
End Function
Public Function ULGetCBlue(Index As Integer) As Integer
    ULGetCBlue = Sprite(Index).UL.CBlue
End Function
Public Function ULGetSRed(Index As Integer) As Integer
    ULGetSRed = Sprite(Index).UL.SRed
End Function
Public Function ULGetSGreen(Index As Integer) As Integer
    ULGetSGreen = Sprite(Index).UL.SGreen
End Function
Public Function ULGetSBlue(Index As Integer) As Integer
    ULGetSBlue = Sprite(Index).UL.SBlue
End Function

Public Function URGetX(Index As Integer) As Single
    URGetX = Sprite(Index).UR.X
End Function
Public Function URGetY(Index As Integer) As Single
    URGetY = Sprite(Index).UR.Y
End Function
Public Function URGetAlpha(Index As Integer) As Integer
    URGetAlpha = Sprite(Index).UR.Alpha
End Function
Public Function URGetCRed(Index As Integer) As Integer
    URGetCRed = Sprite(Index).UR.CRed
End Function
Public Function URGetCGreen(Index As Integer) As Integer
    URGetCGreen = Sprite(Index).UR.CGreen
End Function
Public Function URGetCBlue(Index As Integer) As Integer
    URGetCBlue = Sprite(Index).UR.CBlue
End Function
Public Function URGetSRed(Index As Integer) As Integer
    URGetSRed = Sprite(Index).UR.SRed
End Function
Public Function URGetSGreen(Index As Integer) As Integer
    URGetSGreen = Sprite(Index).UR.SGreen
End Function
Public Function URGetSBlue(Index As Integer) As Integer
    URGetSBlue = Sprite(Index).UR.SBlue
End Function

Public Function LLGetX(Index As Integer) As Single
    LLGetX = Sprite(Index).LL.X
End Function
Public Function LLGetY(Index As Integer) As Single
    LLGetY = Sprite(Index).LL.Y
End Function
Public Function LLGetAlpha(Index As Integer) As Integer
    LLGetAlpha = Sprite(Index).LL.Alpha
End Function
Public Function LLGetCRed(Index As Integer) As Integer
    LLGetCRed = Sprite(Index).LL.CRed
End Function
Public Function LLGetCGreen(Index As Integer) As Integer
    LLGetCGreen = Sprite(Index).LL.CGreen
End Function
Public Function LLGetCBlue(Index As Integer) As Integer
    LLGetCBlue = Sprite(Index).LL.CBlue
End Function
Public Function LLGetSRed(Index As Integer) As Integer
    LLGetSRed = Sprite(Index).LL.SRed
End Function
Public Function LLGetSGreen(Index As Integer) As Integer
    LLGetSGreen = Sprite(Index).LL.SGreen
End Function
Public Function LLGetSBlue(Index As Integer) As Integer
    LLGetSBlue = Sprite(Index).LL.SBlue
End Function

Public Function LRGetX(Index As Integer) As Single
    LRGetX = Sprite(Index).LR.X
End Function
Public Function LRGetY(Index As Integer) As Single
    LRGetY = Sprite(Index).LR.Y
End Function
Public Function LRGetAlpha(Index As Integer) As Integer
    LRGetAlpha = Sprite(Index).LR.Alpha
End Function
Public Function LRGetCRed(Index As Integer) As Integer
    LRGetCRed = Sprite(Index).LR.CRed
End Function
Public Function LRGetCGreen(Index As Integer) As Integer
    LRGetCGreen = Sprite(Index).LR.CGreen
End Function
Public Function LRGetCBlue(Index As Integer) As Integer
    LRGetCBlue = Sprite(Index).LR.CBlue
End Function
Public Function LRGetSRed(Index As Integer) As Integer
    LRGetSRed = Sprite(Index).LR.SRed
End Function
Public Function LRGetSGreen(Index As Integer) As Integer
    LRGetSGreen = Sprite(Index).LR.SGreen
End Function
Public Function LRGetSBlue(Index As Integer) As Integer
    LRGetSBlue = Sprite(Index).LR.SBlue
End Function

Public Function GetZOrder(Index As Integer) As Integer
    GetZOrder = Sprite(Index).ZOrder
End Function
Public Function GetVisible(Index As Integer) As Boolean
    GetVisible = Sprite(Index).Visible
End Function
Public Function GetDrawn(Index As Integer) As Boolean
    GetDrawn = Sprite(Index).Drawn
End Function
Public Function GetDrawnX(Index As Integer) As Single
    GetDrawnX = Sprite(Index).DrawnX
End Function
Public Function GetDrawnY(Index As Integer) As Single
    GetDrawnY = Sprite(Index).DrawnY
End Function

Public Function CounterGetCounter(Index As Integer) As Single
    CounterGetCounter = Sprite(Index).Counter
End Function
Public Function CounterGetCounterIncrement(Index As Integer) As Single
    CounterGetCounterIncrement = Sprite(Index).CounterInc
End Function
Public Function CounterGetCounterAuto(Index As Integer) As Boolean
    CounterGetCounterAuto = Sprite(Index).AutoCounterInc
End Function

Public Sub QMSetPutRectangle(Index As Integer, X As Single, Y As Single, Width As Single, Height As Single)
    With Sprite(Index)
        .UL.X = X
        .UL.Y = Y
        .UR.X = X + Width
        .UR.Y = Y
        .LL.X = X
        .LL.Y = Y + Height
        .LR.X = X + Width
        .LR.Y = Y + Height
    End With
End Sub

Public Sub QMExpand(Index As Integer, Value As Single)
    With Sprite(Index)
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
    With Sprite(Index)
        .UL.X = .UL.X - Value
        .UR.X = .UR.X + Value
        .LL.X = .LL.X - Value
        .LR.X = .LR.X + Value
    End With
End Sub

Public Sub QMEXpandHeight(Index As Integer, Value As Single)
    With Sprite(Index)
        .UL.Y = .UL.Y - Value
        .UR.Y = .UR.Y - Value
        .LL.Y = .LL.Y + Value
        .LR.Y = .LR.Y + Value
    End With
End Sub

Public Sub QMSetGetRectangle(Index As Integer, X As Single, Y As Single, Width As Single, Height As Single)
    With Sprite(Index).Get
        .X = X
        .Y = Y
        .Width = Width
        .Height = Height
    End With
    UpDateTValues Index
End Sub

Public Sub QMSetAlpha(Index As Integer, Alpha As Integer)
    With Sprite(Index)
        .UL.Alpha = Alpha
        .UR.Alpha = Alpha
        .LL.Alpha = Alpha
        .LR.Alpha = Alpha
    End With
    UpdateColor Index
End Sub

Public Sub QMSetCRed(Index As Integer, CRed As Integer)
    With Sprite(Index)
        .UL.CRed = CRed
        .UR.CRed = CRed
        .LL.CRed = CRed
        .LR.CRed = CRed
    End With
    UpdateColor Index
End Sub

Public Sub QMSetCGreen(Index As Integer, CGreen As Integer)
    With Sprite(Index)
        .UL.CGreen = CGreen
        .UR.CGreen = CGreen
        .LL.CGreen = CGreen
        .LR.CGreen = CGreen
    End With
    UpdateColor Index
End Sub

Public Sub QMSetCBlue(Index As Integer, CBlue As Integer)
    With Sprite(Index)
        .UL.CBlue = CBlue
        .UR.CBlue = CBlue
        .LL.CBlue = CBlue
        .LR.CBlue = CBlue
    End With
    UpdateColor Index
End Sub

Public Sub QMSetSRed(Index As Integer, SRed As Integer)
    With Sprite(Index)
        .UL.SRed = SRed
        .UR.SRed = SRed
        .LL.SRed = SRed
        .LR.SRed = SRed
    End With
    UpdateSpecular Index
End Sub

Public Sub QMSetSGreen(Index As Integer, SGreen As Integer)
    With Sprite(Index)
        .UL.SGreen = SGreen
        .UR.SGreen = SGreen
        .LL.SGreen = SGreen
        .LR.SGreen = SGreen
    End With
    UpdateSpecular Index
End Sub

Public Sub QMSetSBlue(Index As Integer, SBlue As Integer)
    With Sprite(Index)
        .UL.SBlue = SBlue
        .UR.SBlue = SBlue
        .LL.SBlue = SBlue
        .LR.SBlue = SBlue
    End With
    UpdateSpecular Index
End Sub

Public Sub QMSetColor(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        With .UL
            .CRed = Red
            .CGreen = Green
            .CBlue = Blue
        End With
        With .UR
            .CRed = Red
            .CGreen = Green
            .CBlue = Blue
        End With
        With .LL
            .CRed = Red
            .CGreen = Green
            .CBlue = Blue
        End With
        With .LR
            .CRed = Red
            .CGreen = Green
            .CBlue = Blue
        End With
    End With
    UpdateColor Index
End Sub

Public Sub QMSetSpecular(Index As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With Sprite(Index)
        With .UL
            .SRed = Red
            .SGreen = Green
            .SBlue = Blue
        End With
        With .UR
            .SRed = Red
            .SGreen = Green
            .SBlue = Blue
        End With
        With .LL
            .SRed = Red
            .SGreen = Green
            .SBlue = Blue
        End With
        With .LR
            .SRed = Red
            .SGreen = Green
            .SBlue = Blue
        End With
    End With
    UpdateSpecular Index
End Sub

Public Sub QMSetPositionToSprite(Index As Integer, SpriteReference As Integer, Optional InheritAngle As Boolean)
    Sprite(Index).XPos = Sprite(SpriteReference).XPos
    Sprite(Index).YPos = Sprite(SpriteReference).YPos
    If InheritAngle = True Then
        Sprite(Index).RAngle = Sprite(SpriteReference).RAngle
    End If
End Sub

Private Sub Class_Initialize()
ReDim Sprite(0) As SpriteInfo
End Sub

Private Sub Class_Terminate()
ReDim Sprite(0) As SpriteInfo
End Sub

Private Sub UpdateColor(Index As Integer)
    With Sprite(Index)
        .C1 = D3DColorARGB(.UL.Alpha, .UL.CRed, .UL.CGreen, .UL.CBlue)
        .C2 = D3DColorARGB(.UR.Alpha, .UR.CRed, .UR.CGreen, .UR.CBlue)
        .C3 = D3DColorARGB(.LL.Alpha, .LL.CRed, .LL.CGreen, .LL.CBlue)
        .C4 = D3DColorARGB(.LR.Alpha, .LR.CRed, .LR.CGreen, .LR.CBlue)
    End With
End Sub

Private Sub UpdateSpecular(Index As Integer)
    With Sprite(Index)
        .S1 = D3DColorARGB(255, .UL.SRed, .UL.SGreen, .UL.SBlue)
        .S2 = D3DColorARGB(255, .UR.SRed, .UR.SGreen, .UR.SBlue)
        .S3 = D3DColorARGB(255, .LL.SRed, .LL.SGreen, .LL.SBlue)
        .S4 = D3DColorARGB(255, .LR.SRed, .LR.SGreen, .LR.SBlue)
    End With
End Sub

Private Sub UpDateTValues(Index As Integer)
Dim TSize As Single
If Sprite(Index).TRef = 0 Then Exit Sub
    TSize = TextureInfo(Sprite(Index).TRef).Size
    With Sprite(Index)
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
