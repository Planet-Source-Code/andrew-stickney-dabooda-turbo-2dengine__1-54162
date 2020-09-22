Attribute VB_Name = "DBGlobals"
Option Explicit

'These Variables are To Store The on or off values of initialized components
    Public DaBoodaDisplayOn As Boolean
    Public DaBoodaSoundOn As Boolean
    Public DaBoodaMusicOn As Boolean
    Public DaBoodaKeyInputOn As Boolean
    Public DaBoodaMouseOn As Boolean
    Public DaBoodaJoyStickOn As Boolean

'These Variables are the core of DirectX8
    Public DirectX As New DirectX8              'Main DirectX
    Public Direct3D As Direct3D8                'Direct3d Core
    Public Direct3DDevice As Direct3DDevice8    'Device to render to
    Public D3DX As New D3DX8                    'Helper Library to createTextures

'Variables for DirectInput
    Public DirectInput As DirectInput8
    Public kDirectInputDevice As DirectInputDevice8
    Public DIState1 As DIKEYBOARDSTATE
    Public DevProp As DIPROPLONG

'Variables for Mouse
    Public mDirectInputDevice As DirectInputDevice8
    Public mDevData() As DIDEVICEOBJECTDATA
    Public mEvents As Long
    Public mDevProp As DIPROPLONG

'Variables for DirectJoyStick
    Public JoyStickDevice As DirectInputDevice8
    Public JoyStickCaps As DIDEVCAPS
    Public JoyStickState As DIJOYSTATE
    Public JoyStickRange As DIPROPRANGE
    Public JoyStickDevEnum As DirectInputEnumDevices8
    
'Variables for DirectShow Objects
    Public DSAudio  As IBasicAudio         'Basic Audio Objectt
    Public DSEvent As IMediaEvent        'MediaEvent Object
    Public DSControl As IMediaControl    'MediaControl Object
    Public DSPosition As IMediaPosition 'MediaPosition Object

'Variables for DirectSound
    Public DirectSound As DirectSound8
    Public DirectSoundEnum As DirectSoundEnum8
    'for Directional Sound
    Public DSMaxVolume As Long
    Public DSFieldSize As Long
    Public DSDecay As Long

'Globals for font
    Public ScreenFont As D3DXFont
    Public ScreenFontDesc As IFont
    Public TextRect As RECT
    Public StandardFont As New StdFont
    
'Globals for simple Rects
    Public PRect As RECT
    Public GRect As RECT

'Variables for Geometry
    'This represents the vertex format we will use
    Public Const FVF_VertexType = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
    'This is your vertex structure,identical to directx7
    Public Type TLVertex
        X As Single
        Y As Single
        z As Single
        rhw As Single
        Color As Long
        specular As Long
        tu As Single
        tv As Single
    End Type

'This is a strip, where your geometry is stored, as a tlvertex type
    'This one strip will be used over and over again as needed
    'all four corners are needed
    Public TextureStrip(0 To 3) As TLVertex
    
'Defined types
'Defined color type
    Public Type DBColor
        Red As Integer
        Green As Integer
        Blue As Integer
        Alpha As Integer
    End Type
'Define Special Vertex Type
    Public Type DBVertex
        X As Single
        Y As Single
        CRed As Integer
        CGreen As Integer
        CBlue As Integer
        SRed As Integer
        SGreen As Integer
        SBlue As Integer
        Alpha As Integer
    End Type
'Define Special Type for Graphic Get
    Public Type DBGet
        X As Single
        Y As Single
        Width As Single
        Height As Single
    End Type
    
'Variables for Screen Mapping and display
    Public DisplayWidth As Single
    Public DisplayHeight As Single
    Public DisplayColor As Long
    Public DisplayColorInfo As DBColor
    
'This Rect Variable is Where the Map and Sprites are actual Drawn
    Public MapView As RECT
    Public MapXCount As Integer
    Public MapYCount As Integer
    Public MapClip As Single

'this Variable Determines how Many Levels are on the screen
'It is is used to determine zorder
    Public MaxLevel As Integer

'For Textures
    Public TextureCount As Integer
    Public TextureColor As DBColor
    Public TextureColorKey As Long
    Public TextureFolder As String
    
'A type to hold information on the textures
    Public Type TextureData
        Size As Single
        Format As CONST_D3DFORMAT
        FileType As Integer
    End Type

'Texture Info
    Public TextureInfo() As TextureData
    
'The textures
    Public Texture() As Direct3DTexture8
    
'The Surfaces for the Textures
    Public Surface() As Direct3DSurface8


'Variables for Overlays
'Type
    Public Type OverlayInfo
        UL As DBVertex
        UR As DBVertex
        LL As DBVertex
        LR As DBVertex
        Get As DBGet
        TextureIndex As Integer
        Visible As Boolean
        RotationAngle As Single
        ZOrder As Integer
        XPos As Single
        YPos As Single
        Tx1 As Single
        Tx2 As Single
        Ty1 As Single
        Ty2 As Single
        C1 As Long
        C2 As Long
        C3 As Long
        C4 As Long
        S1 As Long
        S2 As Long
        S3 As Long
        S4 As Long
        XMirror As Boolean
        YMirror As Boolean
    End Type
'Variables
    Public Overlay() As OverlayInfo
    Public OverlayCount As Integer

'Variables for Text
'Variables for type
    Public Type TextInfo
        fText As String
        ZOrder As Integer
        Visible As Boolean
        Color As DBColor
        XPos As Single
        YPos As Single
        Width As Single
        Height As Single
    End Type
    
'Variables
    Public TextEX() As TextInfo
    Public TextCount As Integer
    
'Variables for maps
'Variable for map type
    Public Type MapInfo
        XCount As Integer
        YCount As Integer
        XPos As Single
        YPos As Single
        XInc As Single
        YInc As Single
        Width As Single
        Height As Single
        XRef As Single
        YRef As Single
        MoveUp As Boolean
        MoveDown As Boolean
        MoveLeft As Boolean
        MoveRight As Boolean
        Looping As Boolean
        Visible As Boolean
        SubMapWidth As Single
        SubMapHeight As Single
        ZOrder As Single
        Get As DBGet
        Tx1 As Single
        Tx2 As Single
        Ty1 As Single
        Ty2 As Single
    End Type
'Variables
    Public Map(8) As MapInfo
    
'Variables for sub maps
'Variable for Sub Map
    Public Type SubMapInfo
        TextureRef As Integer
        ULCol As DBColor
        URCol As DBColor
        LLCol As DBColor
        LRCol As DBColor
        ULSpec As DBColor
        URSpec As DBColor
        LLSpec As DBColor
        LRSpec As DBColor
        Visible As Boolean
        LUp As Boolean
        LDown As Boolean
        LLeft As Boolean
        LRight As Boolean
        C1 As Long
        C2 As Long
        C3 As Long
        C4 As Long
        S1 As Long
        S2 As Long
        S3 As Long
        S4 As Long
    End Type
'Variables
    Public SubMap(8, 64, 64) As SubMapInfo

'Variables for Sprites
'Variable Type for Sprite
    Public Type SpriteInfo
        XPos As Single
        YPos As Single
        Radius As Single
        RAngle As Single
        TRef As Integer
        Sref As Integer
        Mref As Integer
        Get As DBGet
        UL As DBVertex
        UR As DBVertex
        LL As DBVertex
        LR As DBVertex
        ZOrder As Integer
        Visible As Boolean
        Drawn As Boolean
        DrawnX As Single
        DrawnY As Single
        Counter As Single
        CounterInc As Single
        AutoCounterInc As Boolean
        Tx1 As Single
        Tx2 As Single
        Ty1 As Single
        Ty2 As Single
        C1 As Long
        C2 As Long
        C3 As Long
        C4 As Long
        S1 As Long
        S2 As Long
        S3 As Long
        S4 As Long
        TempX As Single
        TempY As Single
        TempAngle As Single
        Void As Boolean
        XMirror As Boolean
        YMirror As Boolean
        InheritAngle As Boolean
    End Type
'Variables
    Public Sprite() As SpriteInfo
    Public SpriteCount As Integer
    
'Variables for AutoMove
'Variable type for automove info
    Public Type AutoMoveInfo
        Width As Single
        Height As Single
        SWidth As Single
        SHeight As Single
        Mref As Single
        Sref As Single
        On As Boolean
        Left As Boolean
        Right As Boolean
        Up As Boolean
        Down As Boolean
    End Type
'Variable
    Public AutoMove As AutoMoveInfo
    Public MapMoved As Boolean
    
'Variables for KeyInput
'Public type for keyinputinfo
    Public Type KeyInputInfo
        Up As Boolean
        Down As Boolean
        Pressed As Boolean
        AutoFire As Boolean
        RKey As Integer
        Repeat As Integer
        RepInc As Integer
    End Type
'Variables
    Public KeyInput() As KeyInputInfo
    Public KeyCount As Integer

'Variable for Music
    Public MusicFolder As String
    Public MusicRepeat As Boolean
    
'Variable for sound
    Public SoundFolder As String
    Public DSBuffer() As DirectSoundSecondaryBuffer8
    Public DSSoundCount As Integer
    
