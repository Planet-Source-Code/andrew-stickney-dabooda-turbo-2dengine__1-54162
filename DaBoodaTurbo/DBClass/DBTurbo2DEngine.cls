VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBTurbo2DEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'These Are Instances of all the classes
Public DBMath As New DBMath
Public DBTexture As New DBTexture
Public DBOverlay As New DBOverlay
Public DBFPS As New DBFPS
Public DBText As New DBText
Public DBMap As New DBMap
Public DBSubMap As New DBSubMap
Public DBSprite As New DBSprite
Public DBAutoMove As New DBAutoMove
Public DBKeyInput As New DBKeyInput
Public DBJoyStick As New DBJoyStick
Public DBMouse As New DBMouse
Public DBMusic As New DBMusic
Public DBSound As New DBSound
Public DBMapData As New DBMapData

Public Sub InitializeDisplay(hwnd&, dWindowed As Boolean, Optional dWidth& = 320, Optional dHeight& = 240, Optional dAdapter As Long = 0, Optional DevType As CONST_D3DDEVTYPE)
    On Error GoTo NotInitialized
    
'Variables for displayMode
    Dim dParams As D3DPRESENT_PARAMETERS
    Dim DispMode As D3DDISPLAYMODE
    
'Set up Screen Variables
    DisplayWidth = dWidth
    DisplayHeight = dHeight
    
'Initialize Direct3d
    
    Set DirectX = New DirectX8
    Set Direct3D = DirectX.Direct3DCreate
    Set D3DX = New D3DX8

'Get Current DisplayMode
    Direct3D.GetAdapterDisplayMode dAdapter, DispMode

'Set the Device to Hardware if empty
    If DevType = 0 Then DevType = D3DDEVTYPE_HAL

'Set Up Parameters for displaymode
    If dWindowed = True Then
        
        With dParams
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = dWidth
        .BackBufferHeight = dHeight
        .EnableAutoDepthStencil = False
        .Windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY
        End With
    Else:
        With dParams
        .BackBufferFormat = DispMode.Format
        .EnableAutoDepthStencil = False
        .BackBufferWidth = dWidth
        .BackBufferHeight = dHeight
        .Windowed = 0
        .SwapEffect = D3DSWAPEFFECT_FLIP
        End With
    End If
    
'Set up DisplayMode
    Set Direct3DDevice = Direct3D.CreateDevice(dAdapter, DevType, hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, dParams)
        
'Set Up Vertex Format
    Direct3DDevice.SetVertexShader FVF_VertexType
    
'Turn Off Lighting
Direct3DDevice.SetRenderState D3DRS_LIGHTING, False
Direct3DDevice.SetRenderState D3DRS_SPECULARENABLE, True

'Set Render State
    Direct3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    Direct3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    Direct3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
'Set Up Texture Stages
    Direct3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    Direct3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
    Direct3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
    DaBoodaDisplayOn = True
    
'Initialize the Classes
Set DBMath = New DBMath
Set DBTexture = New DBTexture
Set DBOverlay = New DBOverlay
Set DBFPS = New DBFPS
Set DBText = New DBText
Set DBMap = New DBMap
Set DBSubMap = New DBSubMap
Set DBSprite = New DBSprite
Set DBAutoMove = New DBAutoMove
Set DBKeyInput = New DBKeyInput
Set DBJoyStick = New DBJoyStick
Set DBMouse = New DBMouse
Set DBMusic = New DBMusic
Set DBSound = New DBSound
Set DBMapData = New DBMapData

    Exit Sub
NotInitialized:
    DaBoodaDisplayOn = False
End Sub

'Subs to set colors
Public Sub SetBackColorRed(Value As Integer)
    DisplayColorInfo.Red = Value
    DisplayColor = D3DColorARGB(255, DisplayColorInfo.Red, DisplayColorInfo.Green, DisplayColorInfo.Blue)
End Sub
Public Sub SetBackColorGreen(Value As Integer)
    DisplayColorInfo.Green = Value
    DisplayColor = D3DColorARGB(255, DisplayColorInfo.Red, DisplayColorInfo.Green, DisplayColorInfo.Blue)
End Sub
Public Sub SetBackColorBlue(Value As Integer)
    DisplayColorInfo.Blue = Value
    DisplayColor = D3DColorARGB(255, DisplayColorInfo.Red, DisplayColorInfo.Green, DisplayColorInfo.Blue)
End Sub
Public Function GetBackColorRed() As Integer
    GetBackColorRed = DisplayColorInfo.Red
End Function
Public Function GetBackColorGreen() As Integer
    GetBackColorGreen = DisplayColorInfo.Green
End Function
Public Function GetBackColorBlue() As Integer
    GetBackColorBlue = DisplayColorInfo.Blue
End Function

'This sets up the view rect, where the maps and sprites are actually drawn
Public Sub SetUpMapView(mLeft As Single, mTop As Single, mRight As Single, mBottom As Single, mClip As Single)
'place Values into mapview rect
With MapView
    .Left = mLeft
    .Right = mRight
    .top = mTop
    .bottom = mBottom
End With
    MapClip = mClip
End Sub

'Subs to Set LevelMax
Public Sub SetMaxLevel(Value As Integer)
    MaxLevel = Value
End Sub
Public Function GetMaxLevel() As Integer
    GetMaxLevel = MaxLevel
End Function

Public Function GetDisplayOn() As Boolean
    GetDisplayOn = DaBoodaDisplayOn
End Function

Public Function GetSoundOn() As Boolean
    GetSoundOn = DaBoodaSoundOn
End Function

Public Function GetMusicOn() As Boolean
    GetMusicOn = DaBoodaMusicOn
End Function

Public Function GetKeyInputOn() As Boolean
    GetKeyInputOn = DaBoodaKeyInputOn
End Function

Public Function GetMouseOn() As Boolean
    GetMouseOn = DaBoodaMouseOn
End Function

Public Function GetJoyStickOn() As Boolean
    GetJoyStickOn = DaBoodaJoyStickOn
End Function

'This Sub is important
'It takes a pixel value for texture and returns it into a value
'from 0 to 1
Private Function TextureValue(Total As Single, Current As Single) As Single
Dim PixelTex As Single
PixelTex = 1 / (Total)
TextureValue = PixelTex * (Current)
End Function

'This Function Makes the Strips
Private Function MakeStrip(X As Single, Y As Single, z As Single, rhw As Single, Color As Long, specular As Long, tu As Single, tv As Single) As TLVertex
MakeStrip.X = X
MakeStrip.Y = Y
MakeStrip.z = z
MakeStrip.rhw = rhw
MakeStrip.Color = Color
MakeStrip.specular = specular
MakeStrip.tu = tu
MakeStrip.tv = tv

End Function

'Sub to render Screen
'Probably the mose important sub in engine
'next to the initialization sub
Public Sub Render()
'Set up Variable for Loop
    Dim Count As Integer
    Dim CountIndex As Integer
    
'Do Screen Limits
    If AutoMove.On = True Then
        CheckAutoMoveSprite
        UpDateMapToLimit
    End If

    CheckMapLimit
    CheckSpritePosition
    CheckSpriteReferences
    
    'clear screen
    Direct3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, DisplayColor, 1#, 0

'Begin scene
    Direct3DDevice.BeginScene

'Start Loop
    For Count = 0 To MaxLevel
'The Graphics
    'Do map
    For CountIndex = 1 To 8
        If Map(CountIndex).ZOrder = Count And Map(CountIndex).Visible = True Then
            RenderMap CountIndex
        End If
    Next CountIndex
    'Do Overlays
If OverlayCount > 0 Then
    For CountIndex = 1 To OverlayCount
        If Overlay(CountIndex).ZOrder = Count Then
            RenderOverlay CountIndex
        End If
    Next CountIndex
End If
    'do Sprites
If SpriteCount > 0 Then
    For CountIndex = 1 To SpriteCount
        If Sprite(CountIndex).AutoCounterInc = True And Sprite(CountIndex).Void = False And Sprite(CountIndex).ZOrder = Count Then
            Sprite(CountIndex).Counter = Sprite(CountIndex).Counter + Sprite(CountIndex).CounterInc
        End If
        If Sprite(CountIndex).ZOrder = Count Then
            RenderSprite CountIndex
        End If
    Next CountIndex
End If
    'Do Text
If TextCount > 0 Then
    For CountIndex = 1 To TextCount
        If TextEX(CountIndex).ZOrder = Count Then
            RenderText CountIndex
        End If
    Next CountIndex
End If
'Increment loop
    Next Count
    
'End Scene
    Direct3DDevice.EndScene

'DoDelay
'do frame delay, to reach FrameRate
    DBFPS.UpdateFPS
    
'Present to display
    Direct3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    
'do events........keeps program from going into infinite loop
    DoEvents

'CheckMusic
If MusicRepeat = True Then
     DBMusic.CheckMusicRepeat
End If

End Sub

'This is the RenderMap Sub
Private Sub RenderMap(Index As Variant)

'Variables for sub
    Dim X1 As Integer, X2 As Single, Y1 As Integer, Y2 As Single     'These are for positions
    Dim TopX As Single, TopY As Single                             'These are the top map positions
    Dim XCount As Integer, YCount As Integer
    Dim XLayer As Integer, YLayer As Integer
    Dim XStart As Single, YStart As Single
    Dim Num As Single
    
    Dim SubMapWidth As Single
    Dim SubMapHeight As Single
    Dim TMap As MapInfo
    Dim TSMap As SubMapInfo
    
    Dim TSize As Single
    
'Put map into Variable
    TMap = Map(Index)
    
'Get MapCounts...How many submaps are viewed on screen
Dim MapXCount As Single, MapYCount As Single
MapXCount = MapView.Right - MapView.Left
MapYCount = MapView.bottom - MapView.top
MapXCount = Int(MapXCount / TMap.SubMapWidth)
MapYCount = Int(MapYCount / TMap.SubMapHeight)

If (MapView.Right - MapView.Left) - (MapXCount * TMap.SubMapWidth) > 0 Then MapXCount = MapXCount + 1
If (MapView.bottom - MapView.top) - (MapYCount * TMap.SubMapHeight) > 0 Then MapYCount = MapYCount + 1

MapXCount = MapXCount + 1
MapYCount = MapYCount + 1

'Find StartPosition
    If TMap.XPos > 0 Then
        X1 = Int(TMap.XPos / TMap.SubMapWidth) + 1
        XLayer = 1 - (X1)
        XStart = TMap.XPos - (TMap.SubMapWidth * X1)
    End If
    If TMap.YPos > 0 Then
        Y1 = Int(TMap.YPos / TMap.SubMapHeight) + 1
        YLayer = 1 - (Y1)
        YStart = TMap.YPos - (TMap.SubMapHeight * Y1)
    End If
    If TMap.XPos < 0 Then
        X1 = Int(Abs(TMap.XPos) / TMap.SubMapWidth)
        XLayer = 1 + X1
        XStart = TMap.XPos + (TMap.SubMapWidth * X1)
    End If
    If TMap.YPos < 0 Then
        Y1 = Int(Abs(TMap.YPos) / TMap.SubMapHeight)
        YLayer = 1 + Y1
        YStart = TMap.YPos + (TMap.SubMapHeight * Y1)
    End If
    If TMap.XPos = 0 Then
        XLayer = 1
        XStart = 0
    End If
    If TMap.YPos = 0 Then
        YLayer = 1
        YStart = 0
    End If
    
'Offset these by mapview Values
            XStart = XStart + MapView.Left
            YStart = YStart + MapView.top

'Loop through all Maps
    For XCount = 1 To MapXCount
        For YCount = 1 To MapYCount
        
'Place position into temp variables
            X1 = XStart + (TMap.SubMapWidth * (XCount - 1))
            Y1 = YStart + (TMap.SubMapHeight * (YCount - 1))

'Set x and y into different variables
    TopX = X1: TopY = Y1
    
'SetLayer Indexes
    X1 = XLayer + (XCount - 1)
    Y1 = YLayer + (YCount - 1)
    If X1 < 1 And TMap.Looping = False Then GoTo DontDraw
    If X1 > TMap.XCount And TMap.Looping = False Then GoTo DontDraw
    If Y1 < 1 And TMap.Looping = False Then GoTo DontDraw
    If Y1 > TMap.YCount And TMap.Looping = False Then GoTo DontDraw
    
    If X1 < 1 Then X1 = TMap.XCount - Abs(X1)
    If X1 > TMap.XCount Then X1 = (X1 - TMap.XCount)
    If Y1 < 1 Then Y1 = TMap.YCount - Abs(Y1)
    If Y1 > TMap.YCount Then Y1 = (Y1 - TMap.YCount)
    
    TSMap = SubMap(Index, X1, Y1)
    If TSMap.Visible = False Then GoTo DontDraw
    
'Set up Texture Size
TSize = TextureInfo(TSMap.TextureRef).Size

'Set up Strip
            TextureStrip(0) = MakeStrip(TopX, TopY, 0, 1, TSMap.C1, TSMap.S1, TMap.Tx1, TMap.Ty1)
            TextureStrip(1) = MakeStrip(TopX + TMap.SubMapWidth, TopY, 0, 1, TSMap.C2, TSMap.S2, TMap.Tx2, TMap.Ty1)
            TextureStrip(2) = MakeStrip(TopX, TopY + TMap.SubMapHeight, 0, 1, TSMap.C3, TSMap.S3, TMap.Tx1, TMap.Ty2)
            TextureStrip(3) = MakeStrip(TopX + TMap.SubMapWidth, TopY + TMap.SubMapHeight, 0, 1, TSMap.C4, TSMap.S4, TMap.Tx2, TMap.Ty2)

'Set texture
            Direct3DDevice.SetTexture 0, Texture(TSMap.TextureRef)
            
'DrawStrip
            Direct3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TextureStrip(0), Len(TextureStrip(0))

'A goto point to goto if layer isn't Drawn
DontDraw:
        Next YCount
    Next XCount

End Sub

Private Sub RenderOverlay(Order As Integer)
Dim tOverlay As OverlayInfo
tOverlay = Overlay(Order)

If tOverlay.Visible = False Then GoTo DontDraw

Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim X3 As Single, Y3 As Single
Dim X4 As Single, Y4 As Single
Dim Angle As Single

Angle = tOverlay.RotationAngle
X1 = tOverlay.XPos + RotateX(tOverlay.UL.X, tOverlay.UL.Y, Angle)
X2 = tOverlay.XPos + RotateX(tOverlay.UR.X, tOverlay.UR.Y, Angle)
X3 = tOverlay.XPos + RotateX(tOverlay.LL.X, tOverlay.LL.Y, Angle)
X4 = tOverlay.XPos + RotateX(tOverlay.LR.X, tOverlay.LR.Y, Angle)

Y1 = tOverlay.YPos + RotateY(tOverlay.UL.X, tOverlay.UL.Y, Angle)
Y2 = tOverlay.YPos + RotateY(tOverlay.UR.X, tOverlay.UR.Y, Angle)
Y3 = tOverlay.YPos + RotateY(tOverlay.LL.X, tOverlay.LL.Y, Angle)
Y4 = tOverlay.YPos + RotateY(tOverlay.LR.X, tOverlay.LR.Y, Angle)

If tOverlay.TextureIndex = 0 Then GoTo DontDraw
TextureStrip(0) = MakeStrip(X1, Y1, 0, 1, tOverlay.C1, tOverlay.S1, tOverlay.Tx1, tOverlay.Ty1)
TextureStrip(1) = MakeStrip(X2, Y2, 0, 1, tOverlay.C2, tOverlay.S2, tOverlay.Tx2, tOverlay.Ty1)
TextureStrip(2) = MakeStrip(X3, Y3, 0, 1, tOverlay.C3, tOverlay.S3, tOverlay.Tx1, tOverlay.Ty2)
TextureStrip(3) = MakeStrip(X4, Y4, 0, 1, tOverlay.C4, tOverlay.S4, tOverlay.Tx2, tOverlay.Ty2)

'Set texture
            Direct3DDevice.SetTexture 0, Texture(tOverlay.TextureIndex)
            
'DrawStrip
            Direct3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TextureStrip(0), Len(TextureStrip(0))

'Exit for not Drawn
DontDraw:
End Sub

Private Sub RenderSprite(Order As Integer)
Dim tSprite As SpriteInfo
tSprite = Sprite(Order)

If tSprite.Void = True Then GoTo DontDraw

Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim X3 As Single, Y3 As Single
Dim X4 As Single, Y4 As Single
Dim Angle As Single

Sprite(Order).Drawn = True
Angle = tSprite.TempAngle
X1 = tSprite.TempX + RotateX(tSprite.UL.X, tSprite.UL.Y, Angle)
X2 = tSprite.TempX + RotateX(tSprite.UR.X, tSprite.UR.Y, Angle)
X3 = tSprite.TempX + RotateX(tSprite.LL.X, tSprite.LL.Y, Angle)
X4 = tSprite.TempX + RotateX(tSprite.LR.X, tSprite.LR.Y, Angle)

Y1 = tSprite.TempY + RotateY(tSprite.UL.X, tSprite.UL.Y, Angle)
Y2 = tSprite.TempY + RotateY(tSprite.UR.X, tSprite.UR.Y, Angle)
Y3 = tSprite.TempY + RotateY(tSprite.LL.X, tSprite.LL.Y, Angle)
Y4 = tSprite.TempY + RotateY(tSprite.LR.X, tSprite.LR.Y, Angle)

Sprite(Order).DrawnX = tSprite.TempX
Sprite(Order).DrawnY = tSprite.TempY
If CheckIntersection(X1, Y1, X2, Y2, X3, Y3, X4, Y4) = False Then
    Sprite(Order).Drawn = False
End If

If Sprite(Order).Drawn = False Or tSprite.Visible = False Then GoTo DontDraw

If tSprite.TRef = 0 Then GoTo DontDraw
TextureStrip(0) = MakeStrip(X1, Y1, 0, 1, tSprite.C1, tSprite.S1, tSprite.Tx1, tSprite.Ty1)
TextureStrip(1) = MakeStrip(X2, Y2, 0, 1, tSprite.C2, tSprite.S2, tSprite.Tx2, tSprite.Ty1)
TextureStrip(2) = MakeStrip(X3, Y3, 0, 1, tSprite.C3, tSprite.S3, tSprite.Tx1, tSprite.Ty2)
TextureStrip(3) = MakeStrip(X4, Y4, 0, 1, tSprite.C4, tSprite.S4, tSprite.Tx2, tSprite.Ty2)

'Set texture
            Direct3DDevice.SetTexture 0, Texture(tSprite.TRef)
            
'DrawStrip
            Direct3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TextureStrip(0), Len(TextureStrip(0))

'Exit for not Drawn
DontDraw:
End Sub

Private Sub RenderText(Order As Integer)
    Dim Color As Long
    Dim TText As TextInfo
    
    TText = TextEX(Order)
    Color = D3DColorARGB(TText.Color.Alpha, TText.Color.Red, TText.Color.Green, TText.Color.Blue)
    
    If TText.Visible = False Then Exit Sub
    
    With TextRect
        .Left = TText.XPos
        .top = TText.YPos
        .Right = .Left + TText.Width
        .bottom = .top + TText.Height
    End With
    
    D3DX.DrawText ScreenFont, Color, TText.fText, TextRect, DT_TOP Or DT_LEFT
    
End Sub

Private Sub CheckAutoMoveSprite()
Dim Sref As Integer, Mref As Integer
Sref = AutoMove.Sref: Mref = Sprite(Sref).Mref

Dim X As Single, Y As Single
X = Sprite(Sref).XPos + Map(Mref).XPos
Y = Sprite(Sref).YPos + Map(Mref).YPos

Dim srect As RECT, CRect As RECT
With srect
    .Left = X - AutoMove.SWidth
    .top = Y - AutoMove.SHeight
    .Right = .Left + (AutoMove.SWidth * 2)
    .bottom = .top + (AutoMove.SHeight * 2)
End With

X = MapView.Right - MapView.Left
X = X / 2
X = X - (AutoMove.Width / 2)
Y = MapView.bottom - MapView.top
Y = Y / 2
Y = Y - (AutoMove.Height / 2)

With CRect
    .Left = X
    .top = Y
    .Right = .Left + AutoMove.Width
    .bottom = .top + AutoMove.Height
End With

    If srect.Left < CRect.Left And AutoMove.Left = True Then MoveMapRight Mref
    If srect.Right > CRect.Right And AutoMove.Right = True Then MoveMapLeft Mref
    If srect.top < CRect.top And AutoMove.Up = True Then MoveMapDown Mref
    If srect.bottom > CRect.bottom And AutoMove.Down = True Then MoveMapUp Mref

End Sub

Private Sub UpDateMapToLimit()
Dim XIndex As Single, YIndex As Single
Dim Mref As Integer, Sref As Integer
Dim Left As Single, Right As Single
Dim top As Single, bottom As Single
Dim XCount As Integer, YCount As Integer

'Find sub map that sprite is on
Sref = AutoMove.Sref
Mref = Sprite(Sref).Mref

XCount = Map(Mref).XCount
YCount = Map(Mref).YCount

XIndex = Sprite(Sref).XPos
YIndex = Sprite(Sref).YPos

XIndex = Int(XIndex / Map(Mref).SubMapWidth) + 1
YIndex = Int(YIndex / Map(Mref).SubMapHeight) + 1

Left = Map(Mref).XPos + ((XIndex - 1) * Map(Mref).SubMapWidth)
top = Map(Mref).YPos + ((YIndex - 1) * Map(Mref).SubMapHeight)
Right = Left + Map(Mref).SubMapWidth
bottom = top + Map(Mref).SubMapHeight

If XIndex < 1 Then XIndex = XCount + XIndex
If XIndex > XCount Then XIndex = XIndex - XCount
If YIndex < 1 Then YIndex = YCount + YIndex
If YIndex > YCount Then YIndex = YIndex - YCount

AutoMove.Left = True
AutoMove.Right = True
AutoMove.Up = True
AutoMove.Down = True

If SubMap(Mref, XIndex, YIndex).LLeft = True And Left > MapView.Left Then
    MoveMapLeft Mref
    AutoMove.Left = False
End If
If SubMap(Mref, XIndex, YIndex).LRight = True And Right < MapView.Right Then
    MoveMapRight Mref
    AutoMove.Right = False
End If
If SubMap(Mref, XIndex, YIndex).LUp = True And top > MapView.top Then
    MoveMapUp Mref
    AutoMove.Up = False
End If
If SubMap(Mref, XIndex, YIndex).LDown = True And bottom < MapView.bottom Then
    MoveMapDown Mref
    AutoMove.Down = False
End If

End Sub

Private Sub CheckMapLimit()
Dim Index As Integer
Dim Diff As Single, Size As Single
    
    For Index = 1 To 8
    If Map(Index).Looping = False Then GoTo DontUpdate
        If Map(Index).XPos >= Map(Index).Width Then
            Size = Int(Map(Index).XPos / Map(Index).SubMapWidth)
            Diff = Map(Index).XPos - (Size * Map(Index).SubMapWidth)
            Map(Index).XPos = Diff
        End If
        If Map(Index).XPos <= -Map(Index).Width Then
            Size = Int(Abs(Map(Index).XPos) / Map(Index).SubMapWidth)
            Diff = Map(Index).XPos + (Size * Map(Index).SubMapWidth)
            Map(Index).XPos = Diff
        End If
        If Map(Index).YPos >= Map(Index).Height Then
            Size = Int(Map(Index).YPos / Map(Index).SubMapHeight)
            Diff = Map(Index).YPos - (Size * Map(Index).SubMapHeight)
            Map(Index).YPos = Diff
        End If
        If Map(Index).YPos <= -Map(Index).Height Then
            Size = Int(Abs(Map(Index).YPos) / Map(Index).SubMapHeight)
            Diff = Map(Index).YPos + (Size * Map(Index).SubMapHeight)
            Map(Index).YPos = Diff
        End If
DontUpdate:
    Next Index

End Sub

Private Sub CheckSpritePosition()
Dim Index As Integer
Dim CRect As RECT
Dim Ref As Variant
Dim X1 As Single, Y1 As Single
    If SpriteCount = 0 Then Exit Sub
    
    For Index = 1 To SpriteCount
        X1 = Sprite(Index).XPos
        Y1 = Sprite(Index).YPos
        Ref = Sprite(Index).Mref
        If Sprite(Index).Sref <> 0 Then GoTo DontOffset
        
    If Ref > 0 And Map(Ref).Looping = True Then
        
        With CRect
            .Left = -(Map(Ref).XPos) - MapClip
            .top = -(Map(Ref).YPos) - MapClip
            .Right = .Left + Map(Ref).Width - 1
            .bottom = .top + Map(Ref).Height - 1
        End With
        
    'Check Farthest Right
        If X1 >= CRect.Left + Map(Ref).Width And X1 <= CRect.Right + Map(Ref).Width Then X1 = X1 - Map(Ref).Width
    'Check Farthest Left
        If X1 >= CRect.Left - Map(Ref).Width And X1 <= CRect.Right - Map(Ref).Width Then X1 = X1 + Map(Ref).Width
    'Check Farthest Down
        If Y1 >= CRect.top + Map(Ref).Height And Y1 <= CRect.bottom + Map(Ref).Height Then Y1 = Y1 - Map(Ref).Height
    'Check Farthest up
        If Y1 >= CRect.top - Map(Ref).Height And Y1 <= CRect.bottom - Map(Ref).Height Then Y1 = Y1 + Map(Ref).Height
    End If

DontOffset:
    Sprite(Index).TempX = X1 + Map(Ref).XPos
    Sprite(Index).TempY = Y1 + Map(Ref).YPos
    Sprite(Index).TempAngle = Sprite(Index).RAngle
    Sprite(Index).XPos = X1
    Sprite(Index).YPos = Y1
    Next Index
End Sub

Private Sub CheckSpriteReferences()
Dim Ref As Integer
Dim Index As Integer
Dim X As Single, Y As Single

If SpriteCount = 0 Then Exit Sub
For Index = 1 To SpriteCount
    Ref = Sprite(Index).Sref
    If Ref <> 0 Then
        X = RotateX(Sprite(Index).XPos, Sprite(Index).YPos, Sprite(Ref).TempAngle)
        Y = RotateY(Sprite(Index).XPos, Sprite(Index).YPos, Sprite(Ref).TempAngle)
        Sprite(Index).TempX = Sprite(Ref).TempX + X
        Sprite(Index).TempY = Sprite(Ref).TempY + Y
        If Sprite(Index).InheritAngle = True Then
            Sprite(Index).TempAngle = Sprite(Ref).TempAngle + Sprite(Index).RAngle
        End If
    End If
Next Index

End Sub

'Rotation Functions
Private Function RotateX(X As Single, Y As Single, Angle As Single) As Single
Dim tx As Single, TAngle As Single
'turns the angle degrees into radians
TAngle = Angle * (3.141592654 / 180)

'Rotate on axis
tx = X * Cos(TAngle) - Y * Sin(TAngle)

RotateX = tx

End Function

Private Function RotateY(X As Single, Y As Single, Angle As Single) As Single
Dim ty As Single, TAngle As Single
'turns the angle degrees into radians
TAngle = Angle * (3.141592654 / 180)

'Rotate on axis
ty = Y * Cos(TAngle) + X * Sin(TAngle)

RotateY = ty

End Function

Private Sub MoveMapUp(Index As Integer)
Dim T As Integer
    Map(Index).YPos = Map(Index).YPos - Map(Index).YInc
    For T = 1 To 8
        If Map(T).YRef = Index Then
            Map(T).YPos = Map(T).YPos - Map(T).YInc
        End If
    Next T
End Sub
Private Sub MoveMapDown(Index As Integer)
Dim T As Integer
    Map(Index).YPos = Map(Index).YPos + Map(Index).YInc
    For T = 1 To 8
        If Map(T).YRef = Index Then
            Map(T).YPos = Map(T).YPos + Map(T).YInc
        End If
    Next T
End Sub
Private Sub MoveMapLeft(Index As Integer)
Dim T As Integer
    Map(Index).XPos = Map(Index).XPos - Map(Index).XInc
    For T = 1 To 8
        If Map(T).XRef = Index Then
            Map(T).XPos = Map(T).XPos - Map(T).XInc
        End If
    Next T
End Sub
Private Sub MoveMapRight(Index As Integer)
Dim T As Integer
    Map(Index).XPos = Map(Index).XPos + Map(Index).XInc
    For T = 1 To 8
        If Map(T).XRef = Index Then
            Map(T).XPos = Map(T).XPos + Map(T).XInc
        End If
    Next T
End Sub

'Private Function CheckIntersection(Index As Integer) As Boolean
'Dim InterSect As Boolean, XPos As Single, YPos As Single
'InterSect = False
'XPos = Sprite(Index).DrawnX
'YPos = Sprite(Index).DrawnY
'If XPos >= MapView.Left - MapClip And XPos <= MapView.Right + MapClip And YPos >= MapView.top - MapClip And YPos <= MapView.bottom + MapClip Then InterSect = True
'
'CheckIntersection = InterSect
'End Function

Private Function CheckIntersection(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, X4 As Single, Y4 As Single) As Boolean
    Dim Check As Boolean
    Check = False
    If PointInMapView(X1, Y1) = True Then Check = True
    If PointInMapView(X2, Y2) = True Then Check = True
    If PointInMapView(X3, Y3) = True Then Check = True
    If PointInMapView(X4, Y4) = True Then Check = True
    CheckIntersection = Check
End Function

Private Function PointInMapView(X As Single, Y As Single) As Boolean
    If X >= MapView.Left And X <= MapView.Right And Y >= MapView.top And Y <= MapView.bottom Then
        PointInMapView = True
    Else:
        PointInMapView = False
    End If
End Function

Private Sub Class_Terminate()
    Set DBTexture = Nothing
    Set DBOverlay = Nothing
    Set DBMath = Nothing
    Set DBFPS = Nothing
    Set DBText = Nothing
    Set DBMap = Nothing
    Set DBSubMap = Nothing
    Set DBSprite = Nothing
    Set DBAutoMove = Nothing
    Set DBKeyInput = Nothing
    Set DBJoyStick = Nothing
    Set DBMusic = Nothing
    Set DBSound = Nothing
    Set DBMapData = Nothing
    
    Set DirectX = Nothing
    Set Direct3D = Nothing
    Set D3DX = Nothing
    
    If DaBoodaKeyInputOn = True Then
        kDirectInputDevice.Unacquire
        Set kDirectInputDevice = Nothing
        Set DirectInput = Nothing
    End If

    If DaBoodaJoyStickOn = True Then
        JoyStickDevice.Unacquire
        Set JoyStickDevice = Nothing
        Set DirectInput = Nothing
    End If
    
    If DaBoodaMouseOn = True Then
        mDirectInputDevice.Unacquire
        Set mDirectInputDevice = Nothing
        Set DirectInput = Nothing
    End If
    
    If DaBoodaMusicOn = True Then
        DSControl.Stop
        Set DSAudio = Nothing
        Set DSEvent = Nothing
        Set DSControl = Nothing
        Set DSPosition = Nothing
    End If

    If DaBoodaSoundOn = True Then
        Set DirectSound = Nothing
        Set DirectSoundEnum = Nothing
    End If

        Set ScreenFontDesc = Nothing
        Set ScreenFont = Nothing

End Sub
