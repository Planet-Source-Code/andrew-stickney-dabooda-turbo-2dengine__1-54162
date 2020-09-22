VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBFPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'variables for Timing Loop
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private curFreq As Currency
Private curStart As Currency
Private curEnd As Currency
Private curFPS As Currency
Private dblResult As Double
Private TempFPS As Single
Private FPS As Single
Private TargetFPS As Single

Public Sub SetFrameRate(Value As Single)
    TargetFPS = Value
    QueryPerformanceFrequency curFreq
End Sub

Public Function GetFPS() As Single
    GetFPS = FPS
End Function

Public Sub UpdateFPS()
'do delay first
Dim bRunning As Boolean
    bRunning = False

Do While bRunning = False
    QueryPerformanceCounter curStart

    If (curStart - curEnd) / curFreq >= 1 / TargetFPS Then
        bRunning = True
    End If

Loop
    TempFPS = TempFPS + 1
    QueryPerformanceCounter curEnd

'check fps
    QueryPerformanceCounter curStart
If (curStart - curFPS) / curFreq >= 1 Then
    FPS = TempFPS
    TempFPS = 0
    QueryPerformanceCounter curFPS
End If

End Sub

