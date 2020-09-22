VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Initialize(hwnd&, Optional Device As Integer = 1)
    'Set up DirectSound
    Set DirectSoundEnum = DirectX.GetDSEnum
    Set DirectSound = DirectX.DirectSoundCreate(DirectSoundEnum.GetGuid(Device))
    DirectSound.SetCooperativeLevel hwnd, DSSCL_NORMAL
    
    DaBoodaSoundOn = True
    Exit Sub
    
NotInitialized:
    DaBoodaSoundOn = False
End Sub

Public Sub SetUpDirectionalSound(MaxVolume As Long, FieldSize As Long, Decay As Long)
    DSMaxVolume = MaxVolume
    DSFieldSize = FieldSize
    DSDecay = Decay
    
    If DSMaxVolume < 0 Then DSMaxVolume = 1
    If DSMaxVolume > 100 Then DSMaxVolume = 100
    
    Dim val As Long
    val = 10000 / 100
    val = val * DSMaxVolume
    DSMaxVolume = -10000 + val
    
End Sub

Public Sub Add(File As String)
    'Add to buffer
    DSSoundCount = DSSoundCount + 1
    ReDim Preserve DSBuffer(DSSoundCount) As DirectSoundSecondaryBuffer8
    
    'Set up BufferDesc
    Dim DirectSoundBufferDesc As DSBUFFERDESC
    
    'Set Desc Flags
    DirectSoundBufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    'Set sound into buffer
    Set DSBuffer(DSSoundCount) = DirectSound.CreateSoundBufferFromFile(SoundFolder + File + ".wav", DirectSoundBufferDesc)
    'set values
    SetVolume DSSoundCount, 100
    SetPan DSSoundCount, 0
End Sub

Public Sub LoadSound(Index As Integer, File As String)
    'Set up BufferDesc
    Dim DirectSoundBufferDesc As DSBUFFERDESC
    
    'Set Desc Flags
    DirectSoundBufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    'Set sound into buffer
    Set DSBuffer(Index) = DirectSound.CreateSoundBufferFromFile(SoundFolder + File + ".wav", DirectSoundBufferDesc)
    'set values
    SetVolume DSSoundCount, 100
    SetPan DSSoundCount, 0
End Sub

Public Sub AddFromSource(SourceIndex As Integer)
    DSSoundCount = DSSoundCount + 1
    ReDim Preserve DSBuffer(DSSoundCount) As DirectSoundSecondaryBuffer8
    
    Set DSBuffer(DSSoundCount) = DSBuffer(SourceIndex)
End Sub

Public Sub Copy(FromIndex As Integer, ToIndex As Integer)
    Set DSBuffer(ToIndex) = DSBuffer(FromIndex)
End Sub

Public Sub Clear()
    ReDim DSBuffer(0) As DirectSoundSecondaryBuffer8
End Sub

Public Sub Remove(Index As Integer)
Dim T As Integer
    If Index <> DSSoundCount Then
        For T = Index To DSSoundCount - 1
        Set DSBuffer(T) = DSBuffer(T + 1)
        Next T
    End If
    DSSoundCount = DSSoundCount - 1
    ReDim Preserve DSBuffer(DSSoundCount) As DirectSoundSecondaryBuffer8
End Sub

Public Sub SetSoundFolder(Path As String)
    SoundFolder = Path
End Sub
Public Function GetSoundFolder() As String
    GetSoundFolder = SoundFolder
End Function

Public Sub PlaySound(Index As Integer, Optional Looping As Boolean)
    'Set buffer to stop and reset position
        DSBuffer(Index).Stop
        DSBuffer(Index).SetCurrentPosition 0
    'play sound looping or not
    If Looping = True Then
        DSBuffer(Index).Play DSBPLAY_LOOPING
    Else:
        DSBuffer(Index).Play DSBPLAY_DEFAULT
    End If
    
End Sub

Public Sub StopSound(Index As Integer)
    'Stop and reset position
        DSBuffer(Index).Stop
        DSBuffer(Index).SetCurrentPosition 0
        
End Sub

Public Sub SetVolume(Index As Integer, Percent As Long)
    'Check Percent
    If Percent < 0 Then Percent = 0
    If Percent > 100 Then Percent = 100
    
    'Set up Percentage to actual Value
    Dim Vol As Long
    Vol = 10000 / 100
    Vol = Percent * Vol
    Vol = -10000 + Vol
    
    'Set Volume
    DSBuffer(Index).SetVolume Vol
    
End Sub

Public Sub SetPan(Index As Integer, Percent As Long)
    'Check Percent
    If Percent < -100 Then Percent = -100
    If Percent > 100 Then Percent = 100
    
    'Set up Percentage to actual value
    Dim Pan As Long
    Pan = 10000 / 100
    Pan = Pan * Percent
    
    'Set Pan
    DSBuffer(Index).SetPan Pan
    
End Sub

Public Sub SetFrequency(Index As Integer, Hrtz As Long)
    'Check Value
    If Hrtz < 100 Then Hrtz = 100
    If Hrtz > 100000 Then Hrtz = 100000
    
    'Set Frequency
    DSBuffer(Index).SetFrequency Hrtz
    
End Sub

Private Sub Class_Initialize()
ReDim DSBuffer(0) As DirectSoundSecondaryBuffer8
End Sub

Private Sub Class_Terminate()
ReDim DSBuffer(0) As DirectSoundSecondaryBuffer8
End Sub
