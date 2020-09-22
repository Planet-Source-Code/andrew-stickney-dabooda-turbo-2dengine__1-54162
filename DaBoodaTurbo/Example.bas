Attribute VB_Name = "Example"
Option Explicit

Public Sub SetUp()
    'Set the engine to new
    Set Engine = New DBTurbo2DEngine
    
    'Init Display... always the first thing to initialize
    SetUpDisplay
    
    'Set up Text
    SetUpText
    
    'Set up Overlay
    SetUpOverlay
    
    'Do a loading screen
    Dim T As Integer
    Engine.DBText.SetText 1, "Loading..."
    Engine.DBText.SetAlpha 1, 0
    For T = 0 To 255 Step 5
        Engine.DBText.SetAlpha 1, T
        Engine.Render
    Next T
    
    'set up levels
    Engine.SetMaxLevel 1

    'Set up textures
    SetUpTextures
    
    'Set up sub Maps
    SetUpMaps
    
    'Set up Glow
    SetUpGlow
    
    'add a sprite
    AddSprite
    
    'set up keyboard
    SetUpKeyBoard
    
    'Set up sound
    SetUpSound
    
    'some last variables
    TargetFPS = 60
    
    'Start the music
    SetUpMusic
    
    'Fade in Example
    For T = 0 To 255 Step 5
        Engine.DBText.SetAlpha 1, 255 - T
        Engine.DBMap.QMSetAlpha 1, T
        Engine.DBMap.QMSetAlpha 2, T
        Engine.DBText.SetText 2, "Frame Rate - " + Format(Engine.DBFPS.GetFPS, "0000") + "/" + Format(TargetFPS, "0000")
        Engine.DBText.SetText 3, "SpriteNum - " + Format(SpriteNum - 1, "000")
        Engine.DBText.SetAlpha 2, T
        Engine.DBText.SetAlpha 3, T
        Engine.DBText.SetAlpha 5, T
        Engine.DBOverlay.QMSetAlpha 2, T
        Engine.Render
    Next T
    
    'do the render loop
    DoRenderLoop
    
End Sub

Public Sub SetUpDisplay()
    'Initialize display
    Engine.InitializeDisplay Main.hwnd, True, 400, 400, D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL
    
    'Set up fps
    Engine.DBFPS.SetFrameRate 60
    
    'Set up map view for mapping system
    Engine.SetUpMapView 0, 0, 400, 400, 256
    
End Sub

Public Sub SetUpSound()
    Engine.DBSound.Initialize Main.hwnd
    With Engine.DBSound
    .SetSoundFolder App.Path + "/Sounds/"
    .Add "Beep1"
    .Add "SpriteHit"
    .Add "SpriteIn"
    .Add "SpriteOut"
    
    End With
    Engine.DBSound.SetUpDirectionalSound 100, 400, 400
    
End Sub

Public Sub SetUpMusic()
    With Engine.DBMusic
        .SetMusicFolder App.Path + "/Music/"
        .LoadMusic "Music"
        .Volume 65
        .PlayMusic True
    End With
End Sub

Public Sub SetUpText()
'Set up some text
    With Engine.DBText
        .CreateFont "timenewroman", 6
        'Text1
        .Add
        .SetXPosition 1, 0
        .SetYPosition 1, 0
        .SetWidth 1, 100
        .SetHeight 1, 16
        .SetVisible 1, True
        'Text2
        .Add
        .SetXPosition 2, 0
        .SetYPosition 2, 0
        .SetWidth 2, 95
        .SetHeight 2, 16
        .SetVisible 2, True
        .QMSetColor 2, 255, 255, 255, 255
        .SetZOrder 2, 1
        'Text3
        .Add
        .SetXPosition 3, 0
        .SetYPosition 3, 10
        .SetWidth 3, 60
        .SetHeight 3, 16
        .SetVisible 3, True
        .QMSetColor 3, 255, 255, 255, 255
        .SetZOrder 3, 1
        'Text4
        .Add
        .SetXPosition 4, 0
        .SetYPosition 4, 0
        .SetWidth 4, 20
        .SetHeight 4, 16
        .SetVisible 4, True
        .QMSetColor 4, 255, 0, 255, 255
        'Text 5
        .AddFromSource 3
        .SetYPosition 5, 20
        .SetText 5, "Push F1 for help..."
    End With
    
End Sub

Public Sub SetUpOverlay()
    With Engine.DBOverlay
        .Add
        .QMSetPutRectangle 1, -150, -150, 300, 300
        .QMSetGetRectangle 1, 0, 0, 256, 256
        .SetTextureReference 1, 78
        .SetVisible 1, False
        .QMSetAlpha 1, 0
        .SetXPosition 1, 200
        .SetYPosition 1, 200
        .SetZOrder 1, 1
        .Add
        .QMSetPutRectangle 2, 0, 0, 128, 128
        .QMSetGetRectangle 2, 0, 0, 128, 128
        .SetTextureReference 2, 79
        .SetVisible 2, True
        .SetXPosition 2, 0
        .SetYPosition 2, 0
        .SetZOrder 2, 1
        
    End With
    
End Sub

Public Sub SetUpTextures()
Dim T As Integer, XT As Single, YT As Single
    'Everything will be done in DBTexture class
    With Engine.DBTexture
    
    'Set texture folder
    .SetFolder App.Path + "/Graphics/"
    .SetColorKeyGreen 255
    
    'Load some textures and do some manipulation, this is just for this example
    'and isn't engine specific... basically I'm loading the sprites and converting them
    'to multiple textures... this is to help me. Speed wise... it doesn't effect anything
    For T = 1 To 64
        .Add 32
    Next T
    .Add 128, "Skull1", , 3
    .Add 128, "Skull2", , 3
    .Add 128, "Skull3", , 3
    .Add 128, "Skull4", , 3
    T = 1
    For YT = 0 To 3
        For XT = 0 To 3
            .CopyRegion 65, T, XT * 32, YT * 32, 0, 0, 32, 32
            T = T + 1
        Next XT
    Next YT
    For YT = 0 To 3
        For XT = 0 To 3
            .CopyRegion 66, T, XT * 32, YT * 32, 0, 0, 32, 32
            T = T + 1
        Next XT
    Next YT
    For YT = 0 To 3
        For XT = 0 To 3
            .CopyRegion 67, T, XT * 32, YT * 32, 0, 0, 32, 32
            T = T + 1
        Next XT
    Next YT
    For YT = 0 To 3
        For XT = 0 To 3
            .CopyRegion 68, T, XT * 32, YT * 32, 0, 0, 32, 32
            T = T + 1
        Next XT
    Next YT
    
    'Load GlowTexture
    .Add 32, "UnderGlow", , 3
    
    'Load Map and Tiles
    .Add 64, "MapBuild", , 3
    .Add 32, "Tile1", , 2
    .Add 32, "Tile2", , 2
    
    'Add blank Maps
    For T = 73 To 77
        .Add 256
    Next T
    
    'Load the help menu
        .Add 256, "HelpMenu", , 3
        .Add 128, "Overlay", , 3
        
    'Make BackMap
    For XT = 0 To 6 Step 2
        For YT = 0 To 6 Step 2
            .CopyRegion 71, 73, 0, 0, XT * 32, YT * 32, 32, 32
        Next YT
    Next XT
        For XT = 1 To 7 Step 2
        For YT = 1 To 7 Step 2
            .CopyRegion 71, 73, 0, 0, XT * 32, YT * 32, 32, 32
        Next YT
    Next XT
    For XT = 1 To 7 Step 2
        For YT = 0 To 6 Step 2
            .CopyRegion 72, 73, 0, 0, XT * 32, YT * 32, 32, 32
        Next YT
    Next XT
        For XT = 0 To 6 Step 2
        For YT = 1 To 7 Step 2
            .CopyRegion 72, 73, 0, 0, XT * 32, YT * 32, 32, 32
        Next YT
    Next XT
    
    'Make up the four corner maps
    Randomize Timer
    Dim R As Integer, XM As Single, YM As Single
    For T = 74 To 77
        For XT = 0 To 15
            For YT = 0 To 15
                R = Int(Rnd * 8)
                If R < 4 Then
                    XM = R * 16
                    YM = 32
                Else:
                    XM = (R - 4) * 16
                    YM = 48
                End If
                .CopyRegion 70, T, XM, YM, XT * 16, YT * 16, 16, 16
            Next YT
        Next XT
    Next T
    
    'Do some side bars
    For T = 0 To 15
        .CopyRegion 70, 74, 0, 16, T * 16, 0, 16, 16
        .CopyRegion 70, 75, 0, 16, T * 16, 0, 16, 16
        .CopyRegion 70, 74, 16, 16, 0, T * 16, 16, 16
        .CopyRegion 70, 76, 16, 16, 0, T * 16, 16, 16
        .CopyRegion 70, 76, 32, 16, T * 16, 240, 16, 16
        .CopyRegion 70, 77, 32, 16, T * 16, 240, 16, 16
        .CopyRegion 70, 75, 48, 16, 240, T * 16, 16, 16
        .CopyRegion 70, 77, 48, 16, 240, T * 16, 16, 16
    Next T

        .CopyRegion 70, 74, 0, 0, 0, 0, 16, 16
        .CopyRegion 70, 75, 16, 0, 240, 0, 16, 16
        .CopyRegion 70, 76, 32, 0, 0, 240, 16, 16
        .CopyRegion 70, 77, 48, 0, 240, 240, 16, 16
        
    'Whew that was alot of work, I am working on a map editor for this engine...
    End With
    
End Sub

Public Sub SetUpMaps()
    Dim T As Integer
    'This is the main sub to set up maps
    With Engine.DBMap
    
    'Map 1
    .SetXCount 1, 2             'Sets the number of sub maps
    .SetYCount 1, 2
    .SetLooping 1, True         'This allows the map to loop over and over
    .SetSubMapHeight 1, 256     'These set how big the submaps are
    .SetSubMapWidth 1, 256
    .SetXIncrement 1, 0.5      'How much to increment with each call to the movemap in the .dbMath
    .SetYIncrement 1, 0.5
    .SetXReference 1, 2         'this allows this map to move with another map, map 2 in this instance
    .SetYReference 1, 2         'this is hand for when you only want to move one map
    .SetVisible 1, True         'This allows us to see the map, default is false
    .QMSetGetRectangle 1, 256, 0, 0, 256, 256 'Allows for you to set where the texture is drawn on the sub maps
    .QMSetAlpha 1, 0            'this is a quick way to set the alpha for every sub map
    .QMSetColor 1, 255, 255, 255 'another quick way to set the color for entire map
    
    'map 2
    .SetXCount 2, 2
    .SetYCount 2, 2
    .SetLooping 2, False
    .SetSubMapHeight 2, 256
    .SetSubMapWidth 2, 256
    .SetVisible 2, True
    .SetXIncrement 2, 1
    .SetYIncrement 2, 1
    .QMSetGetRectangle 2, 256, 0, 0, 256, 256
    .QMSetAlpha 2, 0
    .QMSetColor 2, 255, 255, 255
    
    End With
    
    'Set up the submaps
    With Engine.DBSubMap
    
    'map 1
    .SetTextureReference 1, 1, 1, 73        'These set what texures each sub map will use
    .SetTextureReference 1, 2, 1, 73
    .SetTextureReference 1, 1, 2, 73
    .SetTextureReference 1, 2, 2, 73
        
    'Map 2
    .SetTextureReference 2, 1, 1, 74
    .SetTextureReference 2, 2, 1, 75
    .SetTextureReference 2, 1, 2, 76
    .SetTextureReference 2, 2, 2, 77
    .QMSetLimit 2, 1, 1, True, True, False, False   'these only work when AutoMove is set up
    .QMSetLimit 2, 2, 1, True, False, False, True   'Basically they tell where to stop scrolling and
    .QMSetLimit 2, 1, 2, False, True, True, False   'even automatically move to allow the sub map
    .QMSetLimit 2, 2, 2, False, False, True, True   'to go to a corrosponding edge
    
    End With
    
    'Set up autoMove
    'This allows us to follow a sprite automatically, All you have to do is reference the sprite
    'to follow and everything is automatic... think about it, all you have to do is move the sprite
    'and everything else follows suit automatically.. this is my most favorite class
    With Engine.DBAutoMove
        .SetWidth 48        'These four set commands are to set the area
        .SetHeight 48       'of the screen to adjust to the sprite
        .SetSWidth 16       'it is automatically centered according to these
        .SetSHeight 16      'the closer these are to being the same, the tighter the centering
        .SetMapReference 2  'the map to move
        .SetSpriteReference 1   'the sprite to follow
        .SetOn True             'make sure to turn it on, or hours of head scratching will ensue
    End With
    
    'This next class is a simple data class, Basically it allows you to define
    'integers for each tile on a map, what number you use is up to you, but
    'I am going to use 1 for block and 0 for nothing
    With Engine.DBMapData
        .SetUpMapDataFromMap 2, 16, 16  'Pretty basic really, the first is the map to use, and the last two are the size and width of your tiles
        .SetAll 0                       'reset all the data to 0, really unnecessary, but good practice, for multiple levels in a game
        'now i am going to put a block border around the map
        For T = 1 To 32     'notice I start at one, this is because map data starts there also the 32 is how many tiles accross the map is
            .SetMapData T, 1, 1
            .SetMapData T, 32, 1
            .SetMapData 1, T, 1
            .SetMapData 32, T, 1
        Next T
    End With
    
End Sub

Public Sub SetUpKeyBoard()
    'This sets the keys for the keyboard
    With Engine.DBKeyInput
        .Initialize Main.hwnd
        .Add DIK_UP
        .Add DIK_DOWN
        .Add DIK_LEFT
        .Add DIK_RIGHT
        .Add DIK_ADD
        .Add DIK_SUBTRACT
        .Add DIK_ESCAPE
        .Add DIK_A
        .Add DIK_R
        .Add DIK_F1
        .SetAutoFire 1, True
        .SetAutoFire 2, True
        .SetRepeat 1, 15
        .SetRepeat 2, 15
    End With
End Sub

Public Sub SetUpGlow()
    'this sets up the glow you see under the target sprite
    With Engine.DBSprite
    .Add
    .SetXPosition 1, 32          'Position is where the sprite is, but in this instance, its the offset to the referenced sprite
    .SetYPosition 1, 32
    .SetMapReference 1, 2       'Map coordinates to follow
    .SetVisible 1, True         'lets see it
    .QMSetAlpha 1, 0            'Alpha for all 4 corners
    .QMSetColor 1, 125, 175, 200  'color for the same 4 corners
    .QMSetPutRectangle 1, -16, -16, 32, 32      'The width and height of what is seen, notice the the xposition and yposition is also the origin of rotation
    .QMSetGetRectangle 1, 0, 0, 32, 32          'the get rect for the texture
    .SetTextureReference 1, 69                  'what texture to use
    .SetRadius 1, 8                             'this is for collison detection, all detection is circular
    End With
    GlowDest = 200
    GlowAlpha = 0
    SpriteNum = 1: CheckNum = 1
    TargetSprite = 2
    
End Sub

Public Sub AddSprite()
    'Add number to sprite
    SpriteNum = SpriteNum + 1
    CheckNum = CheckNum + 1
    
    'Create a new sprite
    With Engine.DBSprite
        .Add
        .SetMapReference SpriteNum, 2
        .SetTextureReference SpriteNum, 1
        .SetRadius SpriteNum, 8
        .SetVisible SpriteNum, True
        .QMSetAlpha SpriteNum, 0
        .QMSetPutRectangle SpriteNum, -16, -16, 32, 32
        .QMSetGetRectangle SpriteNum, 0, 0, 32, 32
    End With
    
    Engine.DBSprite.SetXPosition SpriteNum, Engine.DBMath.GetRandomNumber(448) + 48
    Engine.DBSprite.SetYPosition SpriteNum, Engine.DBMath.GetRandomNumber(448) + 48
    Engine.DBSprite.SetRotationAngle SpriteNum, Engine.DBMath.GetRandomNumber(360)
    
    
    SpriteSpec(SpriteNum) = 0
    SpriteSpecDest(SpriteNum) = 0
    SpriteAlpha(SpriteNum) = 0
    SpriteAlphaDest(SpriteNum) = 255
    SpriteAni(SpriteNum) = 1
    
End Sub

Public Sub RemoveSprite1()
    If SpriteNum = 2 Then Exit Sub
    If SpriteAlphaDest(TargetSprite) = 0 Then Exit Sub
    Engine.DBMath.SoundSpritePlay 4, TargetSprite
    SpriteAlphaDest(TargetSprite) = 0
    TargetSprite = TargetSprite - 1
    If TargetSprite < 2 Then TargetSprite = SpriteNum - 1
    If TargetSprite < 2 Then TargetSprite = 2

End Sub

Public Sub RemoveSprite2(Index As Integer)
    Engine.DBSprite.RemoveSprite Index
    
    Dim T As Integer
    For T = Index To SpriteNum
        SpriteSpec(T) = SpriteSpec(T + 1)
        SpriteSpecDest(T) = SpriteSpecDest(T + 1)
        SpriteAlpha(T) = SpriteAlpha(T + 1)
        SpriteAlphaDest(T) = SpriteAlphaDest(T + 1)
        SpriteAni(T) = SpriteAni(T + 1)
    Next T
    
    SpriteNum = SpriteNum - 1
    If TargetSprite > SpriteNum Then TargetSprite = 2
    
    
End Sub

Public Sub DoRenderLoop()
Dim BLooping As Boolean

BLooping = True

'Do the loop
Do While BLooping = True

'check keys
CheckKeyboard

'Update sprites
UpDateSprites

'Update Glow
UpDateGlow

'Update text
Engine.DBText.SetText 2, "Frame Rate - " + Format(Engine.DBFPS.GetFPS, "0000") + "/" + Format(TargetFPS, "0000")
Engine.DBText.SetText 3, "SpriteNum - " + Format(SpriteNum - 1, "000")

'Update FrameRate
Engine.DBFPS.SetFrameRate TargetFPS

'Render Screen
Engine.Render

If Engine.DBKeyInput.ReturnKeyDown(7) = True Then BLooping = False

Loop

End Sub

Public Sub UpDateSprites()
Dim T As Integer, A As Integer
For T = 2 To SpriteNum
    'First move the sprite
    Engine.DBMath.SpriteAngleIncrementUp T, 1
    
    'Check for border hits
    If Engine.DBMapData.GetMapDataBySprite(T, 0, -8) = 1 Then       'Hit top
        Engine.DBMath.SpriteDeflectTopAngle T
        SpriteSpecDest(T) = 170
        Engine.DBMath.SoundSpritePlay 2, T
    End If
    If Engine.DBMapData.GetMapDataBySprite(T, -8, 0) = 1 Then       'Hit left
        Engine.DBMath.SpriteDeflectLeftAngle T
        SpriteSpecDest(T) = 170
        Engine.DBMath.SoundSpritePlay 2, T
    End If
    If Engine.DBMapData.GetMapDataBySprite(T, 8, 0) = 1 Then        'hit right
        Engine.DBMath.SpriteDeflectRightAngle T
        SpriteSpecDest(T) = 170
        Engine.DBMath.SoundSpritePlay 2, T
    End If
    If Engine.DBMapData.GetMapDataBySprite(T, 0, 8) = 1 Then        'hit bottom
        Engine.DBMath.SpriteDeflectBottomAngle T
        SpriteSpecDest(T) = 170
        Engine.DBMath.SoundSpritePlay 2, T
    End If
    
    'Check for collision
    If T = TargetSprite And SpriteNum > 2 Then
        For A = 2 To SpriteNum
            If A <> T Then
                If Engine.DBMath.CheckSpriteCollision(T, A) = True Then
                    Engine.DBMath.FaceSprite T, A
                    Engine.DBMath.FaceSprite A, T
                    Engine.DBMath.SpriteInvertAngle T
                    Engine.DBMath.SpriteInvertAngle A
                    SpriteSpecDest(A) = 170
                    SpriteSpecDest(T) = 170
                    Engine.DBMath.SoundSpritePlay 1, T
                End If
            End If
        Next A
    End If
    
    'update spec
    Engine.DBSprite.QMSetSRed T, SpriteSpec(T)
    If SpriteSpec(T) > SpriteSpecDest(T) Then SpriteSpec(T) = SpriteSpec(T) - 10
    If SpriteSpec(T) < SpriteSpecDest(T) Then SpriteSpec(T) = SpriteSpec(T) + 10
    If SpriteSpec(T) = 170 Then SpriteSpecDest(T) = 0
    
    'update alpha
    Engine.DBSprite.QMSetAlpha T, SpriteAlpha(T)
    If SpriteAlpha(T) > SpriteAlphaDest(T) Then SpriteAlpha(T) = SpriteAlpha(T) - 5
    If SpriteAlpha(T) < SpriteAlphaDest(T) Then SpriteAlpha(T) = SpriteAlpha(T) + 5
    
    'Play sound when added
    If SpriteAlpha(T) = 5 And SpriteAlphaDest(T) = 255 Then Engine.DBMath.SoundSpritePlay 3, T
    
    'update animation
    Engine.DBSprite.SetTextureReference T, Int(SpriteAni(T))
    SpriteAni(T) = SpriteAni(T) + 0.5
    If SpriteAni(T) > 64.5 Then SpriteAni(T) = 1
    
    'update the spec for the target
    If T = TargetSprite Then
        Engine.DBSprite.QMSetSpecular T, 0, 100, 150
    Else:
        Engine.DBSprite.QMSetSpecular T, SpriteSpec(T), 0, 0
    End If
    
Next T

T = 2
Do While T <= SpriteNum
    If SpriteAlpha(T) = 0 Then RemoveSprite2 T
    T = T + 1
Loop

'Set text box
Engine.DBText.SetXPosition 4, Engine.DBSprite.GetDrawnX(TargetSprite) - 16
Engine.DBText.SetYPosition 4, Engine.DBSprite.GetDrawnY(TargetSprite) - 16
Engine.DBText.SetText 4, Str$(TargetSprite - 1)

'check for sprite 2 if spritenum = 2
If SpriteNum = 2 And SpriteAlphaDest(2) = 0 Then
    SpriteAlphaDest(2) = 255
End If

End Sub

Public Sub UpDateGlow()
    If TargetSprite > SpriteNum Then TargetSprite = 2
    If GlowAlpha < GlowDest Then GlowAlpha = GlowAlpha + 5
    If GlowAlpha > GlowDest Then GlowAlpha = GlowAlpha - 5
    If GlowAlpha = 200 And GlowDest = 200 Then GlowDest = 125
    If GlowAlpha = 125 And GlowDest = 125 Then GlowDest = 200
    
    With Engine.DBSprite
        .QMSetAlpha 1, GlowAlpha
    End With
    
    If Engine.DBMath.FindDistance(1, TargetSprite) > 2 Then
    Engine.DBMath.TrackSprite 1, TargetSprite, 2, False
    Else:
    Engine.DBSprite.QMSetPositionToSprite 1, TargetSprite
    End If
    
    
End Sub

Public Sub CheckKeyboard()
    'First we most poll the keys
    Engine.DBKeyInput.PollKeyBoard
    With Engine.DBKeyInput
        If .ReturnKeyDown(1) = True Then TargetFPS = TargetFPS + 5
        If .ReturnKeyDown(2) = True Then TargetFPS = TargetFPS - 5
        If .ReturnKeyDown(3) = True Then TargetSprite = TargetSprite - 1
        If .ReturnKeyDown(4) = True Then TargetSprite = TargetSprite + 1
        If .ReturnKeyDown(5) = True Then AddSprite
        If .ReturnKeyDown(6) = True Then RemoveSprite1
    End With
    If TargetFPS < 20 Then TargetFPS = 20
    If TargetFPS > 9995 Then TargetFPS = 9995
    If TargetSprite < 2 Then TargetSprite = SpriteNum
    If TargetSprite > SpriteNum Then TargetSprite = 2
    Dim T As Integer
    If Engine.DBKeyInput.ReturnKeyDown(8) = True Then
        For T = 2 To SpriteNum
            If T <> TargetSprite Then
                Engine.DBMath.FaceSprite T, TargetSprite
            End If
        Next T
    End If
    If Engine.DBKeyInput.ReturnKeyDown(9) = True Then
        For T = 2 To SpriteNum
            If T <> TargetSprite Then
                Engine.DBMath.FaceSprite T, TargetSprite
                Engine.DBMath.SpriteInvertAngle T
            End If
        Next T
    End If
    If Engine.DBKeyInput.ReturnKeyDown(10) = True Then
        DoHelpMenu
    End If
End Sub

Public Sub DoHelpMenu()
Dim T As Integer
Engine.DBOverlay.SetVisible 1, True
For T = 0 To 255 Step 5
    Engine.DBOverlay.QMSetAlpha 1, T
    Engine.DBOverlay.SetRotationAngle 1, -51 + (T / 5)
    Engine.Render
    Engine.DBText.SetText 2, "Frame Rate - " + Format(Engine.DBFPS.GetFPS, "0000") + "/" + Format(TargetFPS, "0000")
Next T

Dim hlooping As Boolean
hlooping = True
Do While hlooping = True
    Engine.DBKeyInput.PollKeyBoard
    If Engine.DBKeyInput.ReturnKeyDown(10) = True Then hlooping = False
    Engine.Render
    Engine.DBText.SetText 2, "Frame Rate - " + Format(Engine.DBFPS.GetFPS, "0000") + "/" + Format(TargetFPS, "0000")
Loop

For T = 255 To 0 Step -5
    Engine.DBOverlay.QMSetAlpha 1, T
    Engine.DBOverlay.SetRotationAngle 1, 51 - (T / 5)
    Engine.Render
    Engine.DBText.SetText 2, "Frame Rate - " + Format(Engine.DBFPS.GetFPS, "0000") + "/" + Format(TargetFPS, "0000")
Next T
Engine.DBOverlay.SetVisible 1, False
End Sub
