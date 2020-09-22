VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CreateFont(Name As String, Size As Single, Optional Bold As Boolean, Optional Italic As Boolean, Optional Underline As Boolean, Optional StrikeThru As Boolean)
    With StandardFont
        .Name = Name
        .Size = Size
        .Bold = Bold
        .Italic = Italic
        .Underline = Underline
        .Strikethrough = StrikeThru
    End With
    
    Set ScreenFontDesc = StandardFont
    Set ScreenFont = D3DX.CreateFont(Direct3DDevice, ScreenFontDesc.hFont)
    
End Sub

Public Sub Add()
    TextCount = TextCount + 1
    ReDim Preserve TextEX(TextCount) As TextInfo
    With TextEX(TextCount)
        .Color.Alpha = 255
        .Color.Red = 255
        .Color.Green = 255
        .Color.Blue = 255
        .fText = ""
        .Visible = True
        .ZOrder = 0
        .XPos = 0
        .YPos = 0
        .Width = 30
        .Height = 20
    End With
    
End Sub

Public Sub AddFromSource(SourceIndex As Integer)
    TextCount = TextCount + 1
    ReDim Preserve TextEX(TextCount) As TextInfo
    
    TextEX(TextCount) = TextEX(SourceIndex)
End Sub

Public Sub Copy(FromIndex As Integer, ToIndex As Integer)
    TextEX(ToIndex) = TextEX(FromIndex)
End Sub

Public Sub Swap(FromIndex As Integer, ToIndex As Integer)
    TextEX(0) = TextEX(FromIndex)
    TextEX(FromIndex) = TextEX(ToIndex)
    TextEX(ToIndex) = TextEX(0)
End Sub

Public Sub Clear()
    TextCount = 0
    ReDim TextEX(0) As TextInfo
End Sub

Public Sub SetAlpha(Index As Integer, Value As Integer)
    TextEX(Index).Color.Alpha = Value
End Sub
Public Sub SetRed(Index As Integer, Value As Integer)
    TextEX(Index).Color.Red = Value
End Sub
Public Sub SetGreen(Index As Integer, Value As Integer)
    TextEX(Index).Color.Green = Value
End Sub
Public Sub SetBlue(Index As Integer, Value As Integer)
    TextEX(Index).Color.Blue = Value
End Sub

Public Sub SetText(Index As Integer, Value As String)
    TextEX(Index).fText = Value
End Sub
Public Sub SetZOrder(Index As Integer, Value As Integer)
    TextEX(Index).ZOrder = Value
End Sub
Public Sub SetVisible(Index As Integer, Value As Boolean)
    TextEX(Index).Visible = Value
End Sub

Public Sub SetXPosition(Index As Integer, Value As Single)
    TextEX(Index).XPos = Value
End Sub
Public Sub SetYPosition(Index As Integer, Value As Single)
    TextEX(Index).YPos = Value
End Sub
Public Sub SetWidth(Index As Integer, Value As Single)
    TextEX(Index).Width = Value
End Sub
Public Sub SetHeight(Index As Integer, Value As Single)
    TextEX(Index).Height = Value
End Sub


Public Function GetAlpha(Index As Integer) As Integer
    GetAlpha = TextEX(Index).Color.Alpha
End Function
Public Function GetRed(Index As Integer) As Integer
    GetRed = TextEX(Index).Color.Red
End Function
Public Function GetGreen(Index As Integer) As Integer
    GetGreen = TextEX(Index).Color.Green
End Function
Public Function GetBlue(Index As Integer) As Integer
    GetBlue = TextEX(Index).Color.Blue
End Function

Public Function GetText(Index As Integer) As String
    GetText = TextEX(Index).fText
End Function
Public Function GetZOrder(Index As Integer) As Integer
    GetZOrder = TextEX(Index).ZOrder
End Function
Public Function GetVisible(Index As Integer) As Boolean
    GetVisible = TextEX(Index).Visible
End Function

Public Function GetXPosition(Index As Integer) As Single
    GetXPosition = TextEX(Index).XPos
End Function
Public Function GetYPosition(Index As Integer) As Single
    GetYPosition = TextEX(Index).YPos
End Function
Public Function GetWidth(Index As Integer) As Single
    GetWidth = TextEX(Index).Width
End Function
Public Function GetHeight(Index As Integer) As Single
    GetHeight = TextEX(Index).Height
End Function


Public Sub QMSetColor(Index As Integer, Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    With TextEX(Index)
        .Color.Alpha = Alpha
        .Color.Red = Red
        .Color.Green = Green
        .Color.Blue = Blue
    End With
End Sub

Public Sub QMSetPutRectangle(Index As Integer, X As Single, Y As Single, Width As Single, Height As Single)
    With TextEX(Index)
        .XPos = X
        .YPos = Y
        .Width = Width
        .Height = Height
    End With
End Sub

Public Function GetTextCount() As Integer
    GetTextCount = TextCount
End Function

Private Sub Class_Initialize()
    ReDim TextEX(0) As TextInfo
End Sub

Private Sub Class_Terminate()
    ReDim TextEX(0) As TextInfo
End Sub
