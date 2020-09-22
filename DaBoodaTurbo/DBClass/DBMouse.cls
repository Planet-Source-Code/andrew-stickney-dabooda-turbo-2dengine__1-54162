VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MouseOn As Boolean
Private MouseUp As Boolean
Private MouseDown As Boolean
Private MouseLeft As Boolean
Private MouseRight As Boolean
Private mLeftDown As Boolean
Private mLeftAuto As Boolean
Private mLeftPressed As Boolean
Private mRightDown As Boolean
Private mRightAuto As Boolean
Private mRightPressed As Boolean
Private LRepeat As Integer
Private RRepeat As Integer
Private LRInc As Integer
Private RRInc As Integer

'Subs for MouseInput
Public Sub InitializeMouse(hwnd&, Optional Buffer& = 10)
    On Error GoTo NotInitialized
    
'Initialize directinput
    Set DirectInput = DirectX.DirectInputCreate
    Set mDirectInputDevice = DirectInput.CreateDevice("guid_SysMouse")
    Call mDirectInputDevice.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call mDirectInputDevice.SetCooperativeLevel(hwnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)
    
    'Redim DevData
    ReDim mDevData(1 To Buffer) As DIDEVICEOBJECTDATA
    
    'Set up DiProp
    mDevProp.lHow = DIPH_DEVICE
    mDevProp.lObj = 0
    mDevProp.lData = Buffer
    Call mDirectInputDevice.SetProperty("DIPROP_BUFFERSIZE", mDevProp)
    
    'set mouse on
    MouseOn = True
    mLeftAuto = True
    mRightAuto = True
    
    'Aquire Mouse
    mDirectInputDevice.Acquire
    
    DaBoodaMouseOn = True
    Exit Sub
    
NotInitialized:
    DaBoodaMouseOn = False
    
End Sub

Public Sub PollMouse()
    Dim Index As Integer
    mEvents = mDirectInputDevice.GetDeviceData(mDevData, DIGDD_DEFAULT)
    
MouseLeft = False
MouseUp = False
MouseRight = False
MouseDown = False
If mLeftAuto = False And mLeftDown = True And mLeftPressed = True Then
    mLeftDown = False
End If
If mRightAuto = False And mRightDown = True And mRightPressed = True Then
    mRightDown = False
End If

    For Index = 1 To mEvents
        Select Case mDevData(Index).lOfs
            Case DIMOFS_X
                If mDevData(Index).lData < 0 Then MouseLeft = True
                If mDevData(Index).lData > 0 Then MouseRight = True
            Case DIMOFS_Y
                If mDevData(Index).lData < 0 Then MouseUp = True
                If mDevData(Index).lData > 0 Then MouseDown = True
            Case DIMOFS_BUTTON0
            If mLeftAuto = True Then
                If mDevData(Index).lData = 128 Then
                    mLeftDown = True
                    mLeftPressed = True
                End If
                If mDevData(Index).lData = 0 Then
                    mLeftDown = False
                    mLeftPressed = False
                End If
            Else:
                If mDevData(Index).lData = 128 And mLeftAuto = False And mLeftPressed = False Then
                    mLeftDown = True
                    mLeftPressed = True
                End If
                If mDevData(Index).lData = 0 And mLeftAuto = False And mLeftPressed = True Then
                    mLeftDown = False
                    mLeftPressed = False
                End If
            End If
             Case DIMOFS_BUTTON1
             If mRightAuto = True Then
                If mDevData(Index).lData = 128 Then
                    mRightDown = True
                    mRightPressed = True
                End If
                If mDevData(Index).lData = 0 Then
                    mRightDown = False
                    mRightPressed = False
                End If
            Else:
                If mDevData(Index).lData = 128 And mRightPressed = False Then
                    mRightDown = True
                    mRightPressed = True
                End If
                If mDevData(Index).lData = 0 And mRightPressed = True Then
                    mRightDown = False
                    mRightPressed = False
                End If
            End If
        End Select
        
    Next Index
    
If mLeftAuto = True Then
    If mLeftPressed = False Then LRInc = LRepeat
    If mLeftPressed = True And LRInc < LRepeat Then mLeftDown = False
    If mLeftPressed = True And LRInc >= LRepeat Then
        mLeftDown = True
        LRInc = 1
    End If
    If mLeftPressed = True And mLeftDown = False Then LRInc = LRInc + 1
End If
If mRightAuto = True Then
    If mRightPressed = False Then RRInc = RRepeat
    If mRightPressed = True And RRInc < RRepeat Then mRightDown = False
    If mRightPressed = True And RRInc >= RRepeat Then
        mRightDown = True
        RRInc = 1
    End If
    If mRightPressed = True And mRightDown = False Then RRInc = RRInc + 1
End If
End Sub

Public Sub SetLeftAutoFire(Value As Boolean)
    mLeftAuto = Value
End Sub

Public Sub SetRightAutoFire(Value As Boolean)
    mRightAuto = Value
End Sub

Public Function GetMouseLeft() As Boolean
    GetMouseLeft = MouseLeft
End Function

Public Function GetMouseUp() As Boolean
    GetMouseUp = MouseUp
End Function

Public Function GetMouseRight() As Boolean
    GetMouseRight = MouseRight
End Function

Public Function GetMouseDown() As Boolean
    GetMouseDown = MouseDown
End Function

Public Function GetLeftButton() As Boolean
    GetLeftButton = mLeftDown
End Function

Public Function GetRightButton() As Boolean
    GetRightButton = mRightDown
End Function

Public Sub SetLeftRepeat(Value As Integer)
    LRepeat = Value
    LRInc = LRepeat
End Sub
Public Sub SetRightRepeat(Value As Integer)
    RRepeat = Value
    RRInc = RRepeat
End Sub

Public Function GetLeftRepeat() As Integer
    GetLeftRepeat = LRInc
End Function
Public Function GetRightRepeat() As Integer
    GetRightRepeat = RRInc
End Function
