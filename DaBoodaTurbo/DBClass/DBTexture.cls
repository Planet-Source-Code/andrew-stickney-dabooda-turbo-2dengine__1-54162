VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Add(TSize As Single, Optional tFile As String, Optional tFormat As CONST_D3DFORMAT, Optional tFileType As Integer)
    TextureCount = TextureCount + 1
    ReDim Preserve TextureInfo(TextureCount) As TextureData
    ReDim Preserve Texture(TextureCount) As Direct3DTexture8
    ReDim Preserve Surface(TextureCount) As Direct3DSurface8
    
    If tFormat = 0 Then tFormat = D3DFMT_A8R8G8B8
    If tFileType < 1 Or tFileType > 4 Then tFileType = 1

    
    Select Case tFileType
        Case 1
            Set Texture(TextureCount) = D3DX.CreateTexture(Direct3DDevice, TSize, TSize, 0, 0, tFormat, D3DPOOL_MANAGED)
        Case 2
            tFile = TextureFolder & tFile & ".bmp"
            Set Texture(TextureCount) = D3DX.CreateTextureFromFileEx(Direct3DDevice, tFile, TSize, TSize, 0, 0, tFormat, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, TextureColorKey, ByVal 0, ByVal 0)
        Case 3
            tFile = TextureFolder & tFile & ".dds"
            Set Texture(TextureCount) = D3DX.CreateTextureFromFileEx(Direct3DDevice, tFile, TSize, TSize, 0, 0, tFormat, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, TextureColorKey, ByVal 0, ByVal 0)
        Case 4
            tFile = TextureFolder & tFile
            Set Texture(TextureCount) = D3DX.CreateTextureFromFileEx(Direct3DDevice, tFile, TSize, TSize, 0, 0, tFormat, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, TextureColorKey, ByVal 0, ByVal 0)
    End Select
    
    Set Surface(TextureCount) = Texture(TextureCount).GetSurfaceLevel(0)
    
    With TextureInfo(TextureCount)
        .FileType = tFileType
        .Format = tFormat
        .Size = TSize
    End With
    
End Sub

Public Sub Replace(Index As Integer, TSize As Single, Optional tFile As String, Optional tFormat As CONST_D3DFORMAT, Optional tFileType As Integer)
    
    If tFormat = 0 Then tFormat = D3DFMT_A8R8G8B8
    If tFileType < 1 Or tFileType > 4 Then tFileType = 1

    
    Select Case tFileType
        Case 1
            Set Texture(Index) = D3DX.CreateTexture(Direct3DDevice, TSize, TSize, 0, 0, tFormat, D3DPOOL_MANAGED)
        Case 2
            tFile = TextureFolder & tFile & ".bmp"
            Set Texture(Index) = D3DX.CreateTextureFromFileEx(Direct3DDevice, tFile, TSize, TSize, 0, 0, tFormat, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, TextureColorKey, ByVal 0, ByVal 0)
        Case 3
            tFile = TextureFolder & tFile & ".dds"
            Set Texture(Index) = D3DX.CreateTextureFromFileEx(Direct3DDevice, tFile, TSize, TSize, 0, 0, tFormat, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, TextureColorKey, ByVal 0, ByVal 0)
        Case 4
            tFile = TextureFolder & tFile
            Set Texture(Index) = D3DX.CreateTextureFromFileEx(Direct3DDevice, tFile, TSize, TSize, 0, 0, tFormat, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, TextureColorKey, ByVal 0, ByVal 0)
    End Select
    
    Set Surface(Index) = Texture(Index).GetSurfaceLevel(0)
    
    With TextureInfo(Index)
        .FileType = tFileType
        .Format = tFormat
        .Size = TSize
    End With
    
End Sub

Public Sub Clear()
    ReDim TextureInfo(0) As TextureData
    ReDim Texture(0) As Direct3DTexture8
    ReDim Surface(0) As Direct3DSurface8
    TextureCount = 0
End Sub

'Sub to copy rect from one texture to another
Public Sub CopyRegion(SourceIndex As Integer, DestIndex As Integer, GetX As Single, GetY As Single, PutX As Single, PutY As Single, Width As Single, Height As Single)
    
'Set Up Rects
    With GRect
        .Left = GetX
        .top = GetY
        .Right = .Left + Width
        .bottom = .top + Height
    End With
    
    With PRect
        .Left = PutX
        .top = PutY
        .Right = .Left + Width
        .bottom = .top + Height
    End With
    
'Blt from surface to surface
Direct3DDevice.CopyRects Surface(SourceIndex), GRect, 1, Surface(DestIndex), PRect

End Sub

Public Sub SetColorKeyRed(Value As Integer)
    TextureColor.Red = Value
    TextureColorKey = D3DColorARGB(255, TextureColor.Red, TextureColor.Green, TextureColor.Blue)
End Sub
Public Sub SetColorKeyGreen(Value As Integer)
    TextureColor.Green = Value
    TextureColorKey = D3DColorARGB(255, TextureColor.Red, TextureColor.Green, TextureColor.Blue)
End Sub
Public Sub SetColorKeyBlue(Value As Integer)
    TextureColor.Blue = Value
    TextureColorKey = D3DColorARGB(255, TextureColor.Red, TextureColor.Green, TextureColor.Blue)
End Sub
Public Function GetColorKeyRed() As Integer
    GetColorKeyRed = TextureColor.Red
End Function
Public Function GetColorKeyGreen() As Integer
    GetColorKeyGreen = TextureColor.Green
End Function
Public Function GetColorKeyBlue() As Integer
    GetColorKeyBlue = TextureColor.Blue
End Function

Public Sub SetFolder(File As String)
    TextureFolder = File
End Sub
Public Function GetFolder() As String
    GetFolder = TextureFolder
End Function

Public Function GetFormat(Index As Integer) As CONST_D3DFORMAT
    GetFormat = TextureInfo(Index).Format
End Function
Public Function GetFileType(Index As Integer) As Integer
    GetFileType = TextureInfo(Index).FileType
End Function
Public Function GetSize(Index As Integer) As Single
    GetSize = TextureInfo(Index).Size
End Function

Public Function GetTextureCount() As Integer
    GetTextureCount = TextureCount
End Function

Private Sub Class_Initialize()
    'redim variables
    ReDim Texture(0) As Direct3DTexture8
    ReDim Surface(0) As Direct3DSurface8
    ReDim TextureInfo(0) As TextureData
    TextureColorKey = D3DColorARGB(255, 0, 0, 0)
    
End Sub

Private Sub Class_Terminate()
Dim Count As Integer
    For Count = 0 To TextureCount
        Set Texture(Count) = Nothing
        Set Surface(Count) = Nothing
    Next Count
End Sub
