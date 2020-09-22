VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'These subs are just a mishmash of different math type functions.
'Also movement subs are located here

'Sub to DeReference the sprite, basically this causes the sprite to inheret
' a new position according to the reference and sets the reference to nothing
Public Sub DeReferenceSprite(Index As Integer, Visible As Boolean)
    If Sprite(Index).Sref = 0 Then Exit Sub
    Dim X As Single, Y As Single, Ref As Integer
    X = Sprite(Index).XPos
    Y = Sprite(Index).YPos
    Ref = Sprite(Index).Sref
    Sprite(Index).XPos = RotateX(X, Y, Sprite(Ref).RAngle) + Sprite(Ref).XPos
    Sprite(Index).YPos = RotateY(X, Y, Sprite(Ref).RAngle) + Sprite(Ref).YPos
    Sprite(Index).RAngle = Sprite(Index).RAngle + Sprite(Ref).RAngle
    Sprite(Index).Visible = Visible
    Sprite(Index).Sref = 0
End Sub

'Subs to MoveMap
Public Sub MoveMapUp(Index As Integer)
MapMoved = True
Dim T As Integer
    Map(Index).YPos = Map(Index).YPos - Map(Index).YInc
    For T = 1 To 8
        If Map(T).YRef = Index Then
            Map(T).YPos = Map(T).YPos - Map(T).YInc
        End If
    Next T
End Sub
Public Sub MoveMapDown(Index As Integer)
MapMoved = True
Dim T As Integer
    Map(Index).YPos = Map(Index).YPos + Map(Index).YInc
    For T = 1 To 8
        If Map(T).YRef = Index Then
            Map(T).YPos = Map(T).YPos + Map(T).YInc
        End If
    Next T
End Sub
Public Sub MoveMapLeft(Index As Integer)
MapMoved = True
Dim T As Integer
    Map(Index).XPos = Map(Index).XPos - Map(Index).XInc
    For T = 1 To 8
        If Map(T).XRef = Index Then
            Map(T).XPos = Map(T).XPos - Map(T).XInc
        End If
    Next T
End Sub
Public Sub MoveMapRight(Index As Integer)
MapMoved = True
Dim T As Integer
    Map(Index).XPos = Map(Index).XPos + Map(Index).XInc
    For T = 1 To 8
        If Map(T).XRef = Index Then
            Map(T).XPos = Map(T).XPos + Map(T).XInc
        End If
    Next T
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

If Left = MapView.Left Then AutoMove.Left = True
If Right = MapView.Right Then AutoMove.Right = True
If top = MapView.top Then AutoMove.Up = True
If bottom = MapView.bottom Then AutoMove.Down = True

End Sub

'Sub to move the map to the sprite... automove must be set up for this to work
Public Sub SetMaptoSprite()
    Dim tx As Single, ty As Single
    Dim mx As Single, my As Single
    
    tx = AutoMove.Width
    ty = AutoMove.Height
    AutoMove.Width = AutoMove.SWidth * 2
    AutoMove.Height = AutoMove.SHeight * 2
    
    Dim OutLoop As Boolean
    OutLoop = False
    Do While OutLoop = False
        MapMoved = False
        CheckAutoMoveSprite
        UpDateMapToLimit
        If MapMoved = False Then OutLoop = True
    Loop
    AutoMove.Width = tx
    AutoMove.Height = ty
    
End Sub

'Subs to Move Sprite
Public Sub SpriteAngleIncrementUp(Index As Integer, Optional Increment As Single)
    Dim X As Single, Y As Single
    If Increment = 0 Then Increment = 1
    X = RotateX(0, -Increment, Sprite(Index).RAngle)
    Y = RotateY(0, -Increment, Sprite(Index).RAngle)
    Sprite(Index).XPos = Sprite(Index).XPos + X
    Sprite(Index).YPos = Sprite(Index).YPos + Y
End Sub
Public Sub SpriteAngleIncrementDown(Index As Integer, Optional Increment As Single)
    Dim X As Single, Y As Single
    If Increment = 0 Then Increment = 1
    X = RotateX(0, Increment, Sprite(Index).RAngle)
    Y = RotateY(0, Increment, Sprite(Index).RAngle)
    Sprite(Index).XPos = Sprite(Index).XPos + X
    Sprite(Index).YPos = Sprite(Index).YPos + Y
End Sub
Public Sub SpriteAngleIncrementLeft(Index As Integer, Optional Increment As Single)
    Dim X As Single, Y As Single
    If Increment = 0 Then Increment = 1
    X = RotateX(-Increment, 0, Sprite(Index).RAngle)
    Y = RotateY(-Increment, 0, Sprite(Index).RAngle)
    Sprite(Index).XPos = Sprite(Index).XPos + X
    Sprite(Index).YPos = Sprite(Index).YPos + Y
End Sub
Public Sub SpriteAngleIncrementRight(Index As Integer, Optional Increment As Single)
    Dim X As Single, Y As Single
    If Increment = 0 Then Increment = 1
    X = RotateX(Increment, 0, Sprite(Index).RAngle)
    Y = RotateY(Increment, 0, Sprite(Index).RAngle)
    Sprite(Index).XPos = Sprite(Index).XPos + X
    Sprite(Index).YPos = Sprite(Index).YPos + Y
End Sub

Public Sub SpriteIncrementUp(Index As Integer, Optional Increment As Single)
    If Increment = 0 Then Increment = 1
    Sprite(Index).YPos = Sprite(Index).YPos - Increment
End Sub
Public Sub SpriteIncrementDown(Index As Integer, Optional Increment As Single)
    If Increment = 0 Then Increment = 1
    Sprite(Index).YPos = Sprite(Index).YPos + Increment
End Sub
Public Sub SpriteIncrementLeft(Index As Integer, Optional Increment As Single)
    If Increment = 0 Then Increment = 1
    Sprite(Index).XPos = Sprite(Index).XPos - Increment
End Sub
Public Sub SpriteIncrementRight(Index As Integer, Optional Increment As Single)
    If Increment = 0 Then Increment = 1
    Sprite(Index).XPos = Sprite(Index).XPos + Increment
End Sub

Public Sub SpriteInvertAngle(Index As Integer)
    Dim Angle As Single
    If Sprite(Index).RAngle = 0 Then Angle = 180
    If Sprite(Index).RAngle = 180 Then Angle = 0
    If Sprite(Index).RAngle < 180 Then Angle = Sprite(Index).RAngle + 180
    If Sprite(Index).RAngle > 180 Then Angle = Sprite(Index).RAngle - 180
    Sprite(Index).RAngle = Angle
End Sub

Public Sub SpriteDeflectTopAngle(Index As Integer)
    Dim RAngle As Single, Angle As Single
    RAngle = Sprite(Index).RAngle: Angle = RAngle
    If RAngle < 90 Then Angle = 180 - RAngle
    If RAngle > 270 Then Angle = 270 - (RAngle - 270)
    Sprite(Index).RAngle = Angle
End Sub
Public Sub SpriteDeflectRightAngle(Index As Integer)
    Dim RAngle As Single, Angle As Single
    RAngle = Sprite(Index).RAngle: Angle = RAngle
    If RAngle = 90 Then Angle = 270
    If RAngle < 90 And RAngle > 0 Then Angle = 360 - RAngle
    If RAngle > 90 And RAngle < 180 Then Angle = 180 + (180 - RAngle)
    Sprite(Index).RAngle = Angle
End Sub
Public Sub SpriteDeflectBottomAngle(Index As Integer)
    Dim RAngle As Single, Angle As Single
    RAngle = Sprite(Index).RAngle: Angle = RAngle
    If RAngle = 180 Then Angle = 0
    If RAngle > 90 And RAngle < 180 Then Angle = 180 - RAngle
    If RAngle < 270 And RAngle > 180 Then Angle = 360 - (RAngle - 180)
    Sprite(Index).RAngle = Angle
End Sub
Public Sub SpriteDeflectLeftAngle(Index As Integer)
    Dim RAngle As Single, Angle As Single
    RAngle = Sprite(Index).RAngle: Angle = RAngle
    If RAngle = 270 Then Angle = 90
    If RAngle > 180 And RAngle < 270 Then Angle = 180 - (RAngle - 180)
    If RAngle > 270 And RAngle < 360 Then Angle = 360 - RAngle
    Sprite(Index).RAngle = Angle
End Sub

Public Sub SpriteOptimizeAngle(Index As Integer)
    Dim Angle As Single, Result As Single
    Angle = Sprite(Index).RAngle
    Result = Int(Angle / 360)
    Sprite(Index).RAngle = Angle - (Result * 360)
End Sub

'A Cool Sub to set up a get Rect by index, used for animation that is clustered
'on one texture
Public Sub SpriteSetGetByIndex(Index As Integer, X As Single, Y As Single, Width As Single, Height As Single, GWidth As Single, GHeight As Single, Position As Integer)
    Dim MaxWidth As Single
    Dim MaxHeight As Single
    Dim XCount As Single
    Dim YCount As Single
    Dim TSize As Single
    Dim GIndex As Single
    TSize = TextureInfo(Sprite(Index).TRef).Size
    GIndex = Position - 1
    XCount = Width / GWidth
    YCount = Height / GHeight
    Dim gx As Single, gy As Single
    gy = Int(GIndex / YCount)
    gx = GIndex - (gy * YCount)
    With Sprite(Index)
    .Get.X = X + gx * GWidth
    .Get.Y = Y + gy * GHeight
    .Get.Width = GWidth
    .Get.Height = GHeight
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

Public Function CheckSpriteCollision(Sprite1 As Integer, Sprite2 As Integer) As Boolean
Dim Result As Single, Distance As Single
Dim DeltaX As Single, DeltaY As Single
Distance = Sprite(Sprite1).Radius + Sprite(Sprite2).Radius
DeltaX = Sprite(Sprite2).XPos - Sprite(Sprite1).XPos
DeltaY = Sprite(Sprite2).YPos - Sprite(Sprite1).YPos
Result = (DeltaX * DeltaX) + (DeltaY * DeltaY)

If Result <= Distance * Distance Then
    CheckSpriteCollision = True
Else:
    CheckSpriteCollision = False
End If

End Function

Public Function FindDistance(Sprite1 As Integer, Sprite2 As Integer) As Double
    Dim X1 As Long, X2 As Long
    Dim Y1 As Long, Y2 As Long
    X1 = Sprite(Sprite1).XPos: Y1 = Sprite(Sprite1).YPos
    X2 = Sprite(Sprite2).XPos: Y2 = Sprite(Sprite2).YPos
    
    FindDistance = Sqr((X2 - X1) * (X2 - X1) + (Y2 - Y1) * (Y2 - Y1))
    
End Function

Private Function TextureValue(Total As Single, Current As Single) As Single
Dim PixelTex As Single
PixelTex = 1 / (Total)
TextureValue = PixelTex * (Current)
End Function

'Subs For Sound
Public Sub SoundSpriteUpdate(SoundIndex As Integer, SpriteIndex As Integer)
Dim XPos As Single, YPos As Single
XPos = Sprite(SpriteIndex).DrawnX
YPos = Sprite(SpriteIndex).DrawnY
    DSBuffer(SoundIndex).SetPan GetDirectionalPan(XPos)
    DSBuffer(SoundIndex).SetVolume GetVolumeDecay(XPos, YPos)
End Sub

Public Sub SoundSpritePlay(SoundIndex As Integer, SpriteIndex As Integer)
Dim XPos As Single, YPos As Single
XPos = Sprite(SpriteIndex).DrawnX
YPos = Sprite(SpriteIndex).DrawnY

'Set Buffer up
    DSBuffer(SoundIndex).SetPan GetDirectionalPan(XPos)
    DSBuffer(SoundIndex).SetVolume GetVolumeDecay(XPos, YPos)

'Stop Sound
    DSBuffer(SoundIndex).Stop
    DSBuffer(SoundIndex).SetCurrentPosition 0
    
'PlaySound
    DSBuffer(SoundIndex).Play DSBPLAY_DEFAULT
    
End Sub

Private Function GetDirectionalPan(XPos As Single) As Long
Dim HalfPoint As Long
Dim IDP As Long
Dim XDef As Long
Dim FinalPan As Long
If XPos > -1 And XPos < DSFieldSize + 1 Then
HalfPoint = DSFieldSize / 2
IDP = 1000 / HalfPoint
XDef = HalfPoint - XPos
FinalPan = XDef * IDP
FinalPan = FinalPan * -1
End If
If XPos < 0 Then FinalPan = -1000
If XPos > DSFieldSize Then FinalPan = 1000
GetDirectionalPan = FinalPan
End Function

Private Function GetVolumeDecay(XPos As Single, YPos As Single) As Long
Dim Xres As Long, Yres As Long
Dim DecRes As Long
Dim IDV As Long
Dim FinalRes As Long

If XPos > -1 And XPos < DSFieldSize + 1 And YPos > -1 And YPos < DSFieldSize + 1 Then
    GetVolumeDecay = DSMaxVolume
End If

'Check X
If XPos < 0 Then Xres = XPos * -1
If XPos > DSFieldSize Then Xres = XPos - DSFieldSize
If XPos > -1 And XPos < DSFieldSize + 1 Then Xres = 0

'checky
If YPos < 0 Then Yres = YPos * -1
If YPos > DSFieldSize Then Yres = YPos - DSFieldSize
If YPos > -1 And YPos < DSFieldSize + 1 Then Yres = 0

'check to see wich is larger
If Xres > Yres Then
DecRes = Xres
Else
DecRes = Yres
End If
If DecRes > DSDecay Then DecRes = DSDecay

'Find IncrementalDecayVolume
IDV = 5000 / DSDecay
DecRes = DecRes * IDV
FinalRes = DSMaxVolume - DecRes
GetVolumeDecay = FinalRes

End Function


'Rotation Functions
Public Function RotateX(X As Single, Y As Single, Angle As Single) As Single
Dim tx As Single, TAngle As Single
'turns the angle degrees into radians
TAngle = Angle * (3.141592654 / 180)

'Rotate on axis
tx = X * Cos(TAngle) - Y * Sin(TAngle)

RotateX = tx

End Function

Public Function RotateY(X As Single, Y As Single, Angle As Single) As Single
Dim ty As Single, TAngle As Single
'turns the angle degrees into radians
TAngle = Angle * (3.141592654 / 180)

'Rotate on axis
ty = Y * Cos(TAngle) + X * Sin(TAngle)

RotateY = ty

End Function

'This sub has one sprite face another
Public Sub FaceSprite(FromSprite As Integer, ToSprite As Integer)
    Dim Angle As Single
    Angle = RetrieveAngle(Sprite(FromSprite).XPos, Sprite(FromSprite).YPos, Sprite(ToSprite).XPos, Sprite(ToSprite).YPos)
    Sprite(FromSprite).RAngle = Angle
End Sub

'This sub Lets a sprite move towards another sprite, and changes angle if wanted
Public Sub TrackSprite(FromSprite As Integer, ToSprite As Integer, Optional Increment As Single, Optional InheritAngle As Boolean)
    Dim Angle As Single, X As Single, Y As Single
    Angle = RetrieveAngle(Sprite(FromSprite).XPos, Sprite(FromSprite).YPos, Sprite(ToSprite).XPos, Sprite(ToSprite).YPos)
    If Increment = 0 Then Increment = 1
    X = RotateX(0, -Increment, Angle)
    Y = RotateY(0, -Increment, Angle)
    Sprite(FromSprite).XPos = Sprite(FromSprite).XPos + X
    Sprite(FromSprite).YPos = Sprite(FromSprite).YPos + Y
    If InheritAngle = True Then
        Sprite(FromSprite).RAngle = Angle
    End If
End Sub

'this sub Returns an angle according to two positions
Public Function RetrieveAngle(FromX As Single, FromY As Single, ToX As Single, ToY As Single) As Single
    Dim Result As Single
    Dim DiffX As Single, DiffY As Single
    Dim FromPoint As Point, ToPoint As Point
    
    FromPoint.X = FromX: FromPoint.Y = FromY
    ToPoint.X = ToX: ToPoint.Y = ToY
    
    DiffX = ToPoint.X - FromPoint.X
    DiffY = ToPoint.Y - FromPoint.Y
    
    If DiffX = 0 And FromPoint.Y > ToPoint.Y Then
        RetrieveAngle = 0
        Exit Function
    End If
    
    If DiffX = 0 And ToPoint.Y > FromPoint.Y Then
        RetrieveAngle = 180
        Exit Function
    End If
    
    If DiffX = 0 And DiffY = 0 Then
        RetrieveAngle = 0
        Exit Function
    End If
    
    Result = Atn(DiffY / DiffX) * 180 / 3.14
    
    If FromPoint.X > ToPoint.X Then Result = Result - 90
    If FromPoint.X < ToPoint.X Then Result = Result + 90
    
    If Result < 0 Then
        Result = Result + 360
    End If
    
    RetrieveAngle = Result
    
End Function

'Simple Function Returns Random Number
Public Function GetRandomNumber(Random As Single, Optional Offset As Single = 1) As Single
    Randomize Timer
    GetRandomNumber = Int(Rnd * Random) + Offset
End Function

'Simple Returns a number for a percentage
Public Function ReturnPercentage(Number As Single, Percent As Single) As Single
    Percent = Percent * 0.01
    ReturnPercentage = Number * Percent
End Function

Public Function FindPercent(Number As Single, Compare As Single) As Single
    Dim Num As Single
    Num = Number / Compare
    FindPercent = Num * 100
End Function

Public Function FindValueBar(Max As Single, Min As Single, Number As Single, MaxValue As Single)
    Dim Length As Single
    Length = Max - Min
    Length = Length / MaxValue
    Length = Int(Number * Length)
    FindValueBar = Length
End Function

