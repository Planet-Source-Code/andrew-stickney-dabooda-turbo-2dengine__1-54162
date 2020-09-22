Attribute VB_Name = "Globals"
Option Explicit

'This is the main engine core
Public Engine As New DBTurbo2DEngine

'These variables are for the sprites, this is just for this example, has nothing to do
'with the engine.
Public TargetSprite As Integer
Public SpriteNum As Integer
Public CheckNum As Integer
Public SpriteSpec(500) As Integer
Public SpriteSpecDest(500) As Integer
Public SpriteAlpha(500) As Integer
Public SpriteAlphaDest(500) As Integer
Public SpriteAni(500) As Single
Public GlowAlpha As Integer
Public GlowDest As Integer

Public TargetFPS As Single

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    Dim lFlag As Long

    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub


