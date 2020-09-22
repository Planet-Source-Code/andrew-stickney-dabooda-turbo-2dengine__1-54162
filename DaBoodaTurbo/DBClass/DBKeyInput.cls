VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBKeyInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Initialize(hwnd&, Optional Buffer& = 10)
On Error GoTo NotInitialized

'Initialize DirectInput
Set DirectInput = DirectX.DirectInputCreate
Set kDirectInputDevice = DirectInput.CreateDevice("GUID_SysKeyboard")

'Set InputDevice1
kDirectInputDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
kDirectInputDevice.SetCooperativeLevel hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

'SetUp Buffer
DevProp.lHow = DIPH_DEVICE
DevProp.lData = Buffer
kDirectInputDevice.SetProperty DIPROP_BUFFERSIZE, DevProp

'Acquire Device
kDirectInputDevice.Acquire

    DaBoodaKeyInputOn = True
    Exit Sub
    
NotInitialized:
    DaBoodaKeyInputOn = False
    
End Sub

Public Sub PollKeyBoard()

    'Get Key States
    kDirectInputDevice.GetDeviceStateKeyboard DIState1
    
End Sub

Public Sub Add(KeyCode As Integer)
KeyCount = KeyCount + 1
ReDim Preserve KeyInput(KeyCount) As KeyInputInfo

    With KeyInput(KeyCount)
        .RKey = KeyCode
        .AutoFire = False
        .Pressed = False
        .Down = False
        .Up = True
        .Repeat = 0
        .RepInc = 0
    End With
End Sub

Public Sub AddFromSource(SourceIndex As Integer)
    KeyCount = KeyCount + 1
    ReDim Preserve KeyInput(KeyCount) As KeyInputInfo
    
    KeyInput(KeyCount) = KeyInput(SourceIndex)
End Sub

Public Sub Copy(FromIndex As Integer, ToIndex As Integer)
    KeyInput(ToIndex) = KeyInput(FromIndex)
End Sub

Public Sub Swap(FromIndex As Integer, ToIndex As Integer)
    KeyInput(0) = KeyInput(FromIndex)
    KeyInput(FromIndex) = KeyInput(ToIndex)
    KeyInput(ToIndex) = KeyInput(FromIndex)
End Sub

Public Sub Clear()
    ReDim KeyInput(0) As KeyInputInfo
    KeyCount = 0
End Sub

Public Sub SetKeyCode(Index As Integer, KeyCode As Integer)
    KeyInput(Index).RKey = KeyCode
End Sub

Public Sub SetAutoFire(Index As Integer, Value As Boolean)
    KeyInput(Index).AutoFire = Value
End Sub

Public Sub SetRepeat(Index As Integer, Value As Integer)
    KeyInput(Index).Repeat = Value
End Sub

Public Function GetKeyCount() As Integer
    GetKeyCount = KeyCount
End Function

Public Function GetKeyCode(Index As Integer) As Integer
    GetKeyCode = KeyInput(Index).RKey
End Function

Public Function GetAutoFire(Index As Integer) As Boolean
    GetAutoFire = KeyInput(Index).AutoFire
End Function

Public Function GetRepeat(Index As Integer) As Integer
    GetRepeat = KeyInput(Index).Repeat
End Function

Public Function ReturnKeyDown(Index As Integer) As Boolean

    If KeyInput(Index).AutoFire = True Then
    'Reset Repeat if key is up
        If KeyInput(Index).Up = True Then
            KeyInput(Index).RepInc = KeyInput(Index).Repeat
        End If
        
    'Reset keydown if repinc < repeat
        If KeyInput(Index).RepInc < KeyInput(Index).Repeat Then
            KeyInput(Index).Down = False
        End If
    
    'Check KeyDown if AutoFire = true
        If DIState1.Key(KeyInput(Index).RKey) = 128 And KeyInput(Index).RepInc >= KeyInput(Index).Repeat Then
            KeyInput(Index).Down = True
            KeyInput(Index).Up = False
            KeyInput(Index).RepInc = 0
        End If
        
    'Check KeyUp if AutoFire = true
        If DIState1.Key(KeyInput(Index).RKey) = 0 Then
            KeyInput(Index).Down = False
            KeyInput(Index).Up = True
        End If
        
    'Increment Repeat if key is down and autofire is on
        If DIState1.Key(KeyInput(Index).RKey) = 128 Then
            KeyInput(Index).RepInc = KeyInput(Index).RepInc + 1
        End If
    End If
    
    If KeyInput(Index).AutoFire = False Then
    'Check KeyPressed if autoFire is Off
        If KeyInput(Index).Down = True And KeyInput(Index).Pressed = True Then
            KeyInput(Index).Down = False
            KeyInput(Index).Up = True
        End If
        
    'Check KeyDown if AutoFire = false
        If DIState1.Key(KeyInput(Index).RKey) = 128 And KeyInput(Index).Pressed = False Then
            KeyInput(Index).Down = True
            KeyInput(Index).Up = False
            KeyInput(Index).Pressed = True
        End If
        
    'CheckKeyUp if autofire = false
        If DIState1.Key(KeyInput(Index).RKey) = 0 And KeyInput(Index).Pressed = True Then
            KeyInput(Index).Down = False
            KeyInput(Index).Up = True
            KeyInput(Index).Pressed = False
        End If
    End If
    
    'Return value
    ReturnKeyDown = KeyInput(Index).Down
    
End Function

Private Sub Class_Initialize()
    ReDim KeyInput(0) As KeyInputInfo
End Sub

Private Sub Class_Terminate()
    ReDim KeyInput(0) As KeyInputInfo
End Sub

